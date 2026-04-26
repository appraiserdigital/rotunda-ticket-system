import express from "express";
import Stripe from "stripe";
import QRCode from "qrcode";
import { v4 as uuidv4 } from "uuid";
import { createClient } from "@supabase/supabase-js";

const app = express();
const PORT = process.env.PORT || 3000;

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
);

const stripe = new Stripe(process.env.STRIPE_SECRET_KEY);

app.use(express.json());
app.use(express.static('.'));

// 🟢 SUCCESS ROUTE (CREATES TICKET DIRECTLY)
app.get("/success", async (req, res) => {
  const sessionId = req.query.session_id;

  try {
    // 1. Get session from Stripe
    const session = await stripe.checkout.sessions.retrieve(sessionId);

    // 2. Check if ticket already exists (prevents duplicates)
    const { data: existing } = await supabase
      .from("tickets")
      .select("*")
      .eq("stripe_session_id", sessionId)
      .single();

    let ticketId;

    if (existing) {
      ticketId = existing.id;
    } else {
      ticketId = uuidv4();

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
        return res.send("Database error");
      }

      console.log("✅ Ticket created:", ticketId);
    }

    res.redirect(`/ticket/${ticketId}`);

  } catch (err) {
    console.log("❌ STRIPE ERROR:", err.message);
    res.send("Payment verification failed");
  }
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

app.listen(PORT, () => {
  console.log("🚀 Server running");
});
