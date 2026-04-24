import express from "express";
import Stripe from "stripe";
import QRCode from "qrcode";
import { v4 as uuidv4 } from "uuid";

const app = express();
const PORT = 3000;

// 🔑 YOUR STRIPE TEST KEY
const stripe = new Stripe("sk_test_YOUR_KEY");

// Storage (temporary – will later replace with DB)
const tickets = {};
const sessionMap = {};

// 🔴 WEBHOOK (Render-safe, no signature check for now)
app.post("/webhook", express.raw({ type: "application/json" }), async (req, res) => {

  let event;

  try {
    event = JSON.parse(req.body.toString());
  } catch (err) {
    console.log("❌ Webhook parse failed:", err.message);
    return res.sendStatus(400);
  }

  if (event.type === "checkout.session.completed") {
    const session = event.data.object;

    const ticketId = uuidv4();

    const ticketUrl = `https://rotunda-ticket-system.onrender.com/check/${ticketId}`;

    const qr = await QRCode.toDataURL(ticketUrl);

    const ticket = {
      id: ticketId,
      name: session.customer_details?.name || "Guest",
      email: session.customer_details?.email || "",
      status: "VALID",
      qr: qr,
      createdAt: new Date()
    };

    tickets[ticketId] = ticket;

    // 🔗 map Stripe session → ticket
    sessionMap[session.id] = ticketId;

    console.log("✅ Ticket created:", ticketId);
    console.log("🔗 Ticket page:", `https://rotunda-ticket-system.onrender.com/ticket/${ticketId}`);
  }

  res.json({ received: true });
});

// JSON middleware
app.use(express.json());

// Static files (scanner)
app.use(express.static('.'));

// 🟢 SUCCESS ROUTE (reliable)
app.get("/success", (req, res) => {
  const sessionId = req.query.session_id;

  const ticketId = sessionMap[sessionId];

  if (!ticketId) {
    return res.send("❌ Ticket not found");
  }

  res.redirect(`/ticket/${ticketId}`);
});

// 🎟 PROFESSIONAL TICKET PAGE
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

// ✅ CHECK (safe)
app.get("/check/:id", (req, res) => {
  const ticket = tickets[req.params.id];

  if (!ticket) return res.json({ status: "INVALID" });

  if (ticket.status === "USED") {
    return res.json({ status: "ALREADY_USED" });
  }

  return res.json({
    status: "VALID",
    name: ticket.name,
    id: ticket.id
  });
});

// 🔥 USE (confirm entry)
app.get("/use/:id", (req, res) => {
  const ticket = tickets[req.params.id];

  if (!ticket) return res.json({ status: "INVALID" });

  if (ticket.status === "USED") {
    return res.json({ status: "ALREADY_USED" });
  }

  ticket.status = "USED";

  return res.json({
    status: "USED",
    name: ticket.name
  });
});

// ROOT
app.get("/", (req, res) => {
  res.send("🚀 Server running");
});

app.listen(PORT, () => {
  console.log("🚀 Server running on port " + PORT);
});
