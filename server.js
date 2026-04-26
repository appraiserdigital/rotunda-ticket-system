import express from "express";
import Stripe from "stripe";
import QRCode from "qrcode";
import { v4 as uuidv4 } from "uuid";
import { createClient } from "@supabase/supabase-js";

const app = express();
const PORT = process.env.PORT || 3000;

// 🟢 SUPABASE
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
);

// 🟢 STRIPE
const stripe = new Stripe(process.env.STRIPE_SECRET_KEY);

// 🔴 WEBHOOK (MUST BE FIRST - RAW BODY)
app.post("/webhook", express.raw({ type: "application/json" }), async (req, res) => {
  console.log("🔥 WEBHOOK HIT");

  let event;

  try {
    const sig = req.headers["stripe-signature"];

    event = stripe.webhooks.constructEvent(
      req.body,
      sig,
      process.env.STRIPE_WEBHOOK_SECRET
    );
  } catch (err) {
    console.log("❌ Webhook signature error:", err.message);
    return res.sendStatus(400);
  }

  if (event.type === "checkout.session.completed") {
    const session = event.data.object;

    const ticketId = uuidv4();

    const { error } = await supabase
      .from("tickets")
      .insert([{
        id: ticketId,
        stripe_session_id: session.id,
        payment_intent_id: session.payment_intent,
        name: session.customer_details?.name || "Guest",
        email: session.customer_details?.email || "",
        amount: session.amount_total,
        currency: session.currency,
        status: "VALID",
        created_at: new Date()
      }]);

    if (error) {
      console.log("❌ SUPABASE INSERT ERROR:", error);
    } else {
      console.log("✅ Ticket created:", ticketId);
    }
  }

  res.json({ received: true });
});

// 🟢 NORMAL MIDDLEWARE (AFTER WEBHOOK)
app.use(express.json());
app.use(express.static('.'));

// 🟢 SUCCESS ROUTE
app.get("/success", async (req, res) => {
  const sessionId = req.query.session_id;

  let attempts = 0;

  const waitForTicket = async () => {
    while (attempts < 20) {
      const { data } = await supabase
        .from("tickets")
        .select("*")
        .eq("stripe_session_id", sessionId)
        .single();

      if (data) return data.id;

      await new Promise(r => setTimeout(r, 1000));
      attempts++;
    }

    return null;
  };

  const ticketId = await waitForTicket();

  if (!ticketId) {
    return res.send("❌ Ticket not ready, refresh");
  }

  res.redirect(`/ticket/${ticketId}`);
});

// 🎟 TICKET PAGE
app.get("/ticket/:id", async (req, res) => {
  const { data: ticket } = await supabase
    .from("tickets")
    .select("*")
    .eq("id", req.params.id)
    .single();

  if (!ticket) return res.send("❌ Ticket not found");

  const qr = await QRCode.toDataURL(
    `https://rotunda-ticket-system-xym6.onrender.com/check/${ticket.id}`
  );

  res.send(`
  <html>
  <body style="font-family: Arial; text-align:center;">
    <h2>🎟 Church Tour Ticket</h2>
    <p>${ticket.name}</p>
    <img src="${qr}" width="200"/>
    <h3>${ticket.status}</h3>
  </body>
  </html>
  `);
});

// 🟢 CHECK
app.get("/check/:id", async (req, res) => {
  const { data: ticket } = await supabase
    .from("tickets")
    .select("*")
    .eq("id", req.params.id)
    .single();

  if (!ticket) return res.json({ status: "INVALID" });

  if (ticket.status === "USED") {
    return res.json({ status: "ALREADY_USED" });
  }

  res.json({ status: "VALID", name: ticket.name });
});

// 🔵 USE
app.get("/use/:id", async (req, res) => {
  const { data: ticket } = await supabase
    .from("tickets")
    .select("*")
    .eq("id", req.params.id)
    .single();

  if (!ticket) return res.json({ status: "INVALID" });

  if (ticket.status === "USED") {
    return res.json({ status: "ALREADY_USED" });
  }

  await supabase
    .from("tickets")
    .update({
      status: "USED",
      used_at: new Date()
    })
    .eq("id", req.params.id);

  res.json({ status: "USED" });
});

// 🚀 START
app.listen(PORT, () => {
  console.log("🚀 Server running");
});
