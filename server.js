import express from "express";
import bodyParser from "body-parser";
import Stripe from "stripe";
import QRCode from "qrcode";
import { v4 as uuidv4 } from "uuid";

const app = express();
const PORT = 3000;

// 🔑 YOUR STRIPE KEYS (leave as you already set them)
const stripe = new Stripe("sk_test_YOUR_KEY");
const endpointSecret = "whsec_Rr4D1NxsBawF7B4PlXyXpD5fGAFGjZ0F";

// In-memory storage
const tickets = {};

// Stripe webhook needs raw body
app.post("/webhook", bodyParser.raw({ type: "application/json" }), async (req, res) => {
  const sig = req.headers["stripe-signature"];

  let event;

  try {
    event = stripe.webhooks.constructEvent(req.body, sig, endpointSecret);
  } catch (err) {
    console.log("❌ Webhook signature failed:", err.message);
    return res.sendStatus(400);
  }

  if (event.type === "checkout.session.completed") {
    const session = event.data.object;

    const ticketId = uuidv4();

    // 👉 IMPORTANT: QR now uses /check (NOT /verify)
    const ticketUrl = `https://tickets.borgdanastasi.com/check/${ticketId}`;

    const qr = await QRCode.toDataURL(ticketUrl);

    tickets[ticketId] = {
      id: ticketId,
      name: session.customer_details?.name || "Guest",
      email: session.customer_details?.email || "",
      status: "VALID",
      qr: qr,
      createdAt: new Date()
    };

    console.log("✅ Ticket created:", ticketId);
    console.log("🔗 Ticket URL:", ticketUrl);
  }

  res.json({ received: true });
});

// Normal JSON parsing
app.use(express.json());

// Serve static files (scanner.html)
app.use(express.static('.'));

// 🎟 Ticket page (QR display ONLY — no logic)

app.get("/ticket/:id", (req, res) => {
  const ticket = tickets[req.params.id];

  if (!ticket) return res.send("❌ Ticket not found");

  res.send(`
  <html>
  <head>
    <title>Your Ticket</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      body {
        margin: 0;
        font-family: Arial;
        background: #f4f6f9;
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
      }

      .ticket {
        background: white;
        border-radius: 15px;
        padding: 25px;
        width: 320px;
        text-align: center;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
      }

      .title {
        font-size: 20px;
        font-weight: bold;
        margin-bottom: 10px;
      }

      .subtitle {
        color: #666;
        margin-bottom: 20px;
      }

      .qr {
        margin: 20px 0;
      }

      .status {
        font-weight: bold;
        padding: 10px;
        border-radius: 8px;
        margin-top: 10px;
      }

      .valid { background: #e8f5e9; color: #2e7d32; }
      .used { background: #fff3e0; color: #ef6c00; }

      .footer {
        margin-top: 15px;
        font-size: 12px;
        color: #999;
      }
    </style>
  </head>

  <body>

    <div class="ticket">

      <div class="title">🎟 Church Tour Ticket</div>
      <div class="subtitle">The Chapels of Malta</div>

      <p><strong>${ticket.name}</strong></p>

      <div class="qr">
        <img src="${ticket.qr}" width="180"/>
      </div>

      <div class="status ${ticket.status === "VALID" ? "valid" : "used"}">
        ${ticket.status}
      </div>

      <div class="footer">
        Ticket ID: ${ticket.id}
      </div>

    </div>

  </body>
  </html>
  `);
});



// ✅ SAFE CHECK (does NOT consume ticket)
app.get("/check/:id", (req, res) => {
  const ticket = tickets[req.params.id];

  if (!ticket) {
    return res.json({ status: "INVALID" });
  }

  if (ticket.status === "USED") {
    return res.json({ status: "ALREADY_USED" });
  }

  return res.json({
    status: "VALID",
    name: ticket.name,
    id: ticket.id
  });
});

// 🔥 FINAL USE (only when confirmed)
app.get("/use/:id", (req, res) => {
  const ticket = tickets[req.params.id];

  if (!ticket) {
    return res.json({ status: "INVALID" });
  }

  if (ticket.status === "USED") {
    return res.json({ status: "ALREADY_USED" });
  }

  ticket.status = "USED";

  return res.json({
    status: "USED",
    name: ticket.name
  });
});

// Root test
app.get("/", (req, res) => {
  res.send("Server running");
});

app.get("/success", (req, res) => {
  const latestTicket = Object.values(tickets).slice(-1)[0];

  if (!latestTicket) {
    return res.send("❌ No ticket found");
  }

  res.redirect(`/ticket/${latestTicket.id}`);
});


app.listen(PORT, () => {
  console.log("🚀 Server running on http://localhost:" + PORT);
});
