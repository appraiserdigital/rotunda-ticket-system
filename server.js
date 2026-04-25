import express from "express";
import Stripe from "stripe";
import QRCode from "qrcode";
import sqlite3 from "sqlite3";
import { open } from "sqlite";
import { v4 as uuidv4 } from "uuid";

const app = express();
const PORT = process.env.PORT || 3000;

const stripe = new Stripe("sk_test_YOUR_KEY");

// 🟢 DATABASE SETUP
let db;

(async () => {
  db = await open({
    filename: "./tickets.db",
    driver: sqlite3.Database
  });

  await db.exec(`
    CREATE TABLE IF NOT EXISTS tickets (
      id TEXT PRIMARY KEY,
      stripeSessionId TEXT,
      paymentIntentId TEXT,
      name TEXT,
      email TEXT,
      amount INTEGER,
      currency TEXT,
      status TEXT,
      createdAt TEXT,
      usedAt TEXT
    )
  `);

  console.log("✅ Database ready");
})();

// 🔴 WEBHOOK
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

    await db.run(`
      INSERT INTO tickets (
        id,
        stripeSessionId,
        paymentIntentId,
        name,
        email,
        amount,
        currency,
        status,
        createdAt
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `, [
      ticketId,
      session.id,
      session.payment_intent,
      session.customer_details?.name || "Guest",
      session.customer_details?.email || "",
      session.amount_total,
      session.currency,
      "VALID",
      new Date().toISOString()
    ]);

    console.log("✅ Ticket created:", ticketId);
  }

  res.json({ received: true });
});

app.use(express.json());
app.use(express.static('.'));

// 🟢 SUCCESS ROUTE (with DB lookup)
app.get("/success", async (req, res) => {
  const sessionId = req.query.session_id;

  let attempts = 0;

  const waitForTicket = () => {
    return new Promise((resolve) => {
      const interval = setInterval(async () => {

        const ticket = await db.get(
          "SELECT * FROM tickets WHERE stripeSessionId = ?",
          [sessionId]
        );

        if (ticket) {
          clearInterval(interval);
          resolve(ticket.id);
        }

        attempts++;

        if (attempts > 10) {
          clearInterval(interval);
          resolve(null);
        }

      }, 500);
    });
  };

  const ticketId = await waitForTicket();

  if (!ticketId) {
    return res.send("❌ Ticket not ready, refresh");
  }

  res.redirect(`/ticket/${ticketId}`);
});

// 🎟 TICKET PAGE
app.get("/ticket/:id", async (req, res) => {
  const ticket = await db.get(
    "SELECT * FROM tickets WHERE id = ?",
    [req.params.id]
  );

  if (!ticket) return res.send("❌ Ticket not found");

  const qr = await QRCode.toDataURL(
    `https://rotunda-ticket-system.onrender.com/check/${ticket.id}`
  );

  res.send(`
  <html>
  <head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      body { font-family: Arial; background:#f4f6f9; display:flex; justify-content:center; align-items:center; height:100vh; }
      .ticket { background:white; padding:25px; border-radius:15px; width:320px; text-align:center; }
      .status { padding:10px; border-radius:8px; margin-top:10px; }
      .valid { background:#e8f5e9; color:#2e7d32; }
      .used { background:#fff3e0; color:#ef6c00; }
    </style>
  </head>
  <body>
    <div class="ticket">
      <h2>🎟 Church Tour Ticket</h2>
      <p><strong>${ticket.name}</strong></p>
      <img src="${qr}" width="180"/>
      <div class="status ${ticket.status === "VALID" ? "valid" : "used"}">
        ${ticket.status}
      </div>
      <p style="font-size:12px;">ID: ${ticket.id}</p>
    </div>
  </body>
  </html>
  `);
});

// 🟢 CHECK
app.get("/check/:id", async (req, res) => {
  const ticket = await db.get(
    "SELECT * FROM tickets WHERE id = ?",
    [req.params.id]
  );

  if (!ticket) return res.json({ status: "INVALID" });

  if (ticket.status === "USED") {
    return res.json({ status: "ALREADY_USED" });
  }

  res.json({
    status: "VALID",
    name: ticket.name,
    id: ticket.id
  });
});

// 🔵 USE
app.get("/use/:id", async (req, res) => {
  const ticket = await db.get(
    "SELECT * FROM tickets WHERE id = ?",
    [req.params.id]
  );

  if (!ticket) return res.json({ status: "INVALID" });

  if (ticket.status === "USED") {
    return res.json({ status: "ALREADY_USED" });
  }

  await db.run(
    "UPDATE tickets SET status = ?, usedAt = ? WHERE id = ?",
    ["USED", new Date().toISOString(), req.params.id]
  );

  res.json({ status: "USED" });
});

app.listen(PORT, () => {
  console.log("🚀 Server running");
});
