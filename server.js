import express from "express";
import Stripe  from "stripe";
import QRCode  from "qrcode";
import { v4 as uuidv4 }    from "uuid";
import { createClient }     from "@supabase/supabase-js";

// ─────────────────────────────────────────────────────────────
// STARTUP — validate every required env var before anything runs
// ─────────────────────────────────────────────────────────────
const REQUIRED_ENV = [
  "SUPABASE_URL",       // https://xxxx.supabase.co
  "SUPABASE_KEY",       // service_role key (NOT anon)
  "STRIPE_SECRET_KEY",  // sk_live_... or sk_test_...
  "BASE_URL",           // https://rotunda-ticket-system-xym6.onrender.com  (no trailing slash)
  "ADMIN_TOKEN",        // any strong random string — protects /admin/* routes
];

for (const key of REQUIRED_ENV) {
  if (!process.env[key]) {
    console.error(`❌ STARTUP FAILED — missing env var: ${key}`);
    process.exit(1);
  }
}

const BASE_URL = process.env.BASE_URL.replace(/\/$/, "");

const app  = express();
const PORT = process.env.PORT || 3000;

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
);

const stripe = new Stripe(process.env.STRIPE_SECRET_KEY);

app.use(express.json());
app.use(express.static("."));

// ─────────────────────────────────────────────────────────────
// HEALTH CHECK
// ─────────────────────────────────────────────────────────────
app.get("/health", (_req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

// ─────────────────────────────────────────────────────────────
// ADMIN AUTH MIDDLEWARE
// Pass token as ?token=xxx OR header X-Admin-Token: xxx
// ─────────────────────────────────────────────────────────────
function requireAdmin(req, res, next) {
  const token =
    req.query.token ||
    req.headers["x-admin-token"] ||
    "";

  if (!token || token !== process.env.ADMIN_TOKEN) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  next();
}

// ─────────────────────────────────────────────────────────────
// SUCCESS — Stripe redirects here after payment
//
// Stripe Payment Link → Success URL must be set to:
//   https://YOUR_DOMAIN/success?session_id={CHECKOUT_SESSION_ID}
// (paste {CHECKOUT_SESSION_ID} literally — Stripe fills it in)
// ─────────────────────────────────────────────────────────────
app.get("/success", async (req, res) => {
  const sessionId = req.query.session_id;

  if (!sessionId) {
    console.error("❌ /success — no session_id in query string");
    return res.status(400).send("Missing session_id parameter.");
  }

  console.log(`🔍 /success — session: ${sessionId}`);

  try {
    // ── 1. Idempotency check ───────────────────────────────
    // maybeSingle() returns null (not error) when no row found
    const { data: existing, error: lookupErr } = await supabase
      .from("tickets")
      .select("id")
      .eq("stripe_session_id", sessionId)
      .maybeSingle();

    if (lookupErr) {
      console.error("❌ Supabase lookup error:", lookupErr);
      return res.status(500).send("Database error during lookup.");
    }

    if (existing) {
      console.log(`ℹ️  Ticket already exists: ${existing.id} — redirecting`);
      return res.redirect(`/ticket/${existing.id}`);
    }

    // ── 2. Retrieve session from Stripe ───────────────────
    let session;
    try {
      session = await stripe.checkout.sessions.retrieve(sessionId);
    } catch (stripeErr) {
      console.error("❌ Stripe retrieve failed:", stripeErr.message);
      return res.status(502).send(`Stripe error: ${stripeErr.message}`);
    }

    // ── 3. Payment must be confirmed ──────────────────────
    //  payment_status === 'paid'  → safe to issue ticket
    //  anything else              → money not captured, no ticket
    if (session.payment_status !== "paid") {
      console.warn(`⚠️  Session ${sessionId} — status: ${session.payment_status}`);
      return res
        .status(402)
        .send(`Payment not completed (status: ${session.payment_status}).`);
    }

    // ── 4. Build the ticket record ────────────────────────
    //  Every financial field is taken directly from Stripe.
    //  We never accept user-supplied amounts.
    const ticketId = uuidv4();
    const now      = new Date().toISOString();

    const record = {
      id:                ticketId,
      stripe_session_id: session.id,                          // idempotency key
      payment_intent_id: session.payment_intent || null,      // for refund look-ups
      name:              session.customer_details?.name  || "Guest",
      email:             session.customer_details?.email || "",
      amount:            session.amount_total,                 // smallest unit (cents)
      currency:          (session.currency || "eur").toLowerCase(),
      payment_status:    session.payment_status,              // 'paid'
      status:            "VALID",
      created_at:        now,
      used_at:           null,
    };

    // ── 5. Insert (atomic — either all columns or nothing) ─
    const { error: insertErr } = await supabase
      .from("tickets")
      .insert([record]);

    if (insertErr) {
      // Postgres unique violation (code 23505) = race condition: another
      // request already inserted. Fetch existing and redirect safely.
      if (insertErr.code === "23505") {
        console.warn("⚠️  Race-condition duplicate — fetching existing ticket");
        const { data: race } = await supabase
          .from("tickets")
          .select("id")
          .eq("stripe_session_id", sessionId)
          .single();
        if (race) return res.redirect(`/ticket/${race.id}`);
      }

      console.error("❌ Supabase insert error:", insertErr);
      return res.status(500).send("Database error creating ticket.");
    }

    console.log(
      `✅ Ticket created: ${ticketId} | ${record.email} | ` +
      `${record.currency.toUpperCase()} ${(record.amount / 100).toFixed(2)}`
    );
    return res.redirect(`/ticket/${ticketId}`);

  } catch (err) {
    console.error("❌ Unexpected error in /success:", err);
    return res.status(500).send("Unexpected server error.");
  }
});

// ─────────────────────────────────────────────────────────────
// TICKET PAGE
// ─────────────────────────────────────────────────────────────
app.get("/ticket/:id", async (req, res) => {
  try {
    const { data: ticket, error } = await supabase
      .from("tickets")
      .select("*")
      .eq("id", req.params.id)
      .single();

    if (error || !ticket) {
      return res.status(404).send("Ticket not found.");
    }

    const checkUrl  = `${BASE_URL}/check/${ticket.id}`;
    const qrDataUrl = await QRCode.toDataURL(checkUrl, {
      errorCorrectionLevel: "H",
      margin: 2,
      width: 300,
    });

    const statusColour = ticket.status === "VALID" ? "#22c55e" : "#ef4444";
    const amount = ticket.amount != null
      ? (ticket.amount / 100).toLocaleString("en-MT", {
          style:    "currency",
          currency: (ticket.currency || "EUR").toUpperCase(),
        })
      : "";
    const dateStr = new Date(ticket.created_at).toLocaleDateString("en-MT", {
      day: "2-digit", month: "long", year: "numeric",
    });

    return res.send(`<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Your Ticket — Rotunda of Mosta</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: Arial, sans-serif;
      background: #0f0f1a;
      color: #e5e5e5;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .ticket {
      background: #1a1a2e;
      border: 1px solid #2a2a4a;
      border-radius: 20px;
      padding: 32px 24px;
      max-width: 380px;
      width: 100%;
      text-align: center;
      box-shadow: 0 25px 60px rgba(0,0,0,0.6);
    }
    .title   { font-size: 22px; font-weight: bold; color: #fff; margin-bottom: 4px; }
    .venue   { font-size: 13px; color: #888; margin-bottom: 24px; letter-spacing: 1px; text-transform: uppercase; }
    .qr-wrap { background: #fff; border-radius: 12px; padding: 14px; display: inline-block; margin-bottom: 22px; }
    .qr-wrap img { display: block; width: 200px; height: 200px; }
    .name    { font-size: 20px; font-weight: bold; color: #fff; margin-bottom: 6px; }
    .detail  { font-size: 13px; color: #999; margin-bottom: 3px; }
    .status  {
      display: inline-block;
      margin-top: 18px;
      padding: 6px 22px;
      border-radius: 20px;
      font-size: 13px;
      font-weight: bold;
      letter-spacing: 1px;
      color: ${statusColour};
      border: 1px solid ${statusColour};
      background: ${statusColour}22;
    }
    .footer { margin-top: 20px; font-size: 11px; color: #444; }
  </style>
</head>
<body>
  <div class="ticket">
    <p class="title">🎟 Church Tour Ticket</p>
    <p class="venue">Rotunda of Mosta</p>
    <div class="qr-wrap"><img src="${qrDataUrl}" alt="Ticket QR Code"/></div>
    <p class="name">${escapeHtml(ticket.name)}</p>
    <p class="detail">${escapeHtml(ticket.email)}</p>
    <p class="detail">${dateStr}${amount ? " · " + amount : ""}</p>
    <span class="status" id="status-badge">${ticket.status}</span>
    <p class="footer" id="footer-text">Present this QR code at the entrance</p>
  </div>

  <script>
    // Poll /check/:id every 4 seconds and update the badge if status changes
    const ticketId  = "${ticket.id}";
    const checkUrl  = "${BASE_URL}/check/" + ticketId;
    let   lastStatus = "${ticket.status}";
    const badge      = document.getElementById("status-badge");
    const footer     = document.getElementById("footer-text");

    function applyStatus(status) {
      if (status === "VALID") {
        badge.style.color      = "#22c55e";
        badge.style.border     = "1px solid #22c55e";
        badge.style.background = "#22c55e22";
        badge.textContent      = "VALID";
        footer.textContent     = "Present this QR code at the entrance";
      } else if (status === "USED" || status === "ALREADY_USED") {
        badge.style.color      = "#ef4444";
        badge.style.border     = "1px solid #ef4444";
        badge.style.background = "#ef444422";
        badge.textContent      = "USED";
        footer.textContent     = "This ticket has already been scanned at the entrance";
        clearInterval(poll); // stop polling once used
      }
    }

    const poll = setInterval(function() {
      fetch(checkUrl)
        .then(r => r.json())
        .then(data => {
          if (data.status !== lastStatus) {
            lastStatus = data.status;
            applyStatus(data.status);
          }
        })
        .catch(function() {}); // silent fail — offline etc.
    }, 4000);
  </script>
</body>
</html>`);

  } catch (err) {
    console.error("❌ Error in /ticket:", err);
    return res.status(500).send("Error loading ticket.");
  }
});

// ─────────────────────────────────────────────────────────────
// CHECK — scanner reads QR, calls this first (read-only)
// Returns: VALID | ALREADY_USED | INVALID
// ─────────────────────────────────────────────────────────────
app.get("/check/:id", async (req, res) => {
  try {
    const { data: ticket, error } = await supabase
      .from("tickets")
      .select("id, status, name")
      .eq("id", req.params.id)
      .single();

    if (error || !ticket) return res.json({ status: "INVALID" });
    if (ticket.status === "USED")
      return res.json({ status: "ALREADY_USED", name: ticket.name });

    return res.json({ status: "VALID", name: ticket.name });

  } catch (err) {
    console.error("❌ Error in /check:", err);
    return res.json({ status: "ERROR" });
  }
});

// ─────────────────────────────────────────────────────────────
// USE — scanner calls this to mark ticket used (write)
// Returns: USED | ALREADY_USED | INVALID
// ─────────────────────────────────────────────────────────────
app.get("/use/:id", async (req, res) => {
  try {
    const { data: ticket, error } = await supabase
      .from("tickets")
      .select("id, status, name")
      .eq("id", req.params.id)
      .single();

    if (error || !ticket) return res.json({ status: "INVALID" });
    if (ticket.status === "USED")
      return res.json({ status: "ALREADY_USED", name: ticket.name });

    const { error: updateErr } = await supabase
      .from("tickets")
      .update({ status: "USED", used_at: new Date().toISOString() })
      .eq("id", req.params.id);

    if (updateErr) {
      console.error("❌ Error marking used:", updateErr);
      return res.json({ status: "ERROR" });
    }

    console.log(`✅ Ticket used: ${req.params.id} (${ticket.name})`);
    return res.json({ status: "USED", name: ticket.name });

  } catch (err) {
    console.error("❌ Error in /use:", err);
    return res.json({ status: "ERROR" });
  }
});

// ═════════════════════════════════════════════════════════════
// ADMIN / FINANCIAL REPORTING ROUTES  (all require ADMIN_TOKEN)
// ═════════════════════════════════════════════════════════════

// ── Summary ──────────────────────────────────────────────────
// GET /admin/summary?token=xxx
// Optional filters: ?from=2025-01-01&to=2025-12-31
app.get("/admin/summary", requireAdmin, async (req, res) => {
  try {
    let q = supabase
      .from("tickets")
      .select("amount, currency, status, created_at");

    if (req.query.from) q = q.gte("created_at", req.query.from);
    if (req.query.to)   q = q.lte("created_at", toEndOfDay(req.query.to));

    const { data: tickets, error } = await q;
    if (error) return res.status(500).json({ error: error.message });

    const total   = tickets.length;
    const used    = tickets.filter(t => t.status === "USED").length;
    const valid   = tickets.filter(t => t.status === "VALID").length;

    // Group revenue by currency (handles multi-currency safely)
    const byCurrency = {};
    for (const t of tickets) {
      const cur = (t.currency || "eur").toUpperCase();
      if (!byCurrency[cur]) byCurrency[cur] = { count: 0, amount_cents: 0 };
      byCurrency[cur].count++;
      byCurrency[cur].amount_cents += t.amount || 0;
    }

    // Add human-readable major unit amounts
    for (const cur of Object.keys(byCurrency)) {
      byCurrency[cur].amount_major =
        (byCurrency[cur].amount_cents / 100).toFixed(2);
    }

    res.json({
      total_tickets:  total,
      used_tickets:   used,
      valid_tickets:  valid,
      by_currency:    byCurrency,
      generated_at:   new Date().toISOString(),
    });

  } catch (err) {
    console.error("❌ /admin/summary error:", err);
    res.status(500).json({ error: "Server error" });
  }
});

// ── Daily breakdown ───────────────────────────────────────────
// GET /admin/daily?token=xxx&from=2025-01-01&to=2025-12-31
app.get("/admin/daily", requireAdmin, async (req, res) => {
  try {
    let q = supabase
      .from("tickets")
      .select("amount, currency, status, created_at")
      .order("created_at", { ascending: true });

    if (req.query.from) q = q.gte("created_at", req.query.from);
    if (req.query.to)   q = q.lte("created_at", toEndOfDay(req.query.to));

    const { data: tickets, error } = await q;
    if (error) return res.status(500).json({ error: error.message });

    // Group by date string (YYYY-MM-DD)
    const byDate = {};
    for (const t of tickets) {
      const date = t.created_at.slice(0, 10);
      if (!byDate[date]) {
        byDate[date] = {
          date,
          count:        0,
          amount_cents: 0,
          currency:     (t.currency || "eur").toUpperCase(),
        };
      }
      byDate[date].count++;
      byDate[date].amount_cents += t.amount || 0;
    }

    const rows = Object.values(byDate).map(d => ({
      ...d,
      amount_major: (d.amount_cents / 100).toFixed(2),
    }));

    res.json(rows);

  } catch (err) {
    console.error("❌ /admin/daily error:", err);
    res.status(500).json({ error: "Server error" });
  }
});

// ── Full ticket list ──────────────────────────────────────────
// GET /admin/tickets?token=xxx
// Optional: &status=VALID|USED  &from=...  &to=...  &email=...
app.get("/admin/tickets", requireAdmin, async (req, res) => {
  try {
    let q = supabase
      .from("tickets")
      .select("*")
      .order("created_at", { ascending: false })
      .limit(1000);   // safety cap — use date filters for large exports

    if (req.query.status) q = q.eq("status", req.query.status);
    if (req.query.from)   q = q.gte("created_at", req.query.from);
    if (req.query.to)     q = q.lte("created_at", toEndOfDay(req.query.to));
    if (req.query.email)  q = q.ilike("email", `%${req.query.email}%`);

    const { data, error } = await q;
    if (error) return res.status(500).json({ error: error.message });

    res.json(data);

  } catch (err) {
    console.error("❌ /admin/tickets error:", err);
    res.status(500).json({ error: "Server error" });
  }
});

// ── CSV export ────────────────────────────────────────────────
// GET /admin/export.csv?token=xxx&from=...&to=...
// Downloads a CSV suitable for Xero / QuickBooks / Excel import
app.get("/admin/export.csv", requireAdmin, async (req, res) => {
  try {
    let q = supabase
      .from("tickets")
      .select("*")
      .order("created_at", { ascending: true });

    if (req.query.status) q = q.eq("status", req.query.status);
    if (req.query.from)   q = q.gte("created_at", req.query.from);
    if (req.query.to)     q = q.lte("created_at", toEndOfDay(req.query.to));

    const { data: tickets, error } = await q;
    if (error) return res.status(500).json({ error: error.message });

    // Columns — ordered for accounting-software import compatibility
    const HEADERS = [
      "ticket_id",
      "date",
      "time",
      "name",
      "email",
      "amount",          // major unit (e.g. 5.00) — what Xero expects
      "currency",
      "ticket_status",
      "payment_status",
      "payment_intent_id",
      "stripe_session_id",
      "used_at",
    ];

    const rows = tickets.map(t => {
      const dt = new Date(t.created_at);
      return [
        csvCell(t.id),
        dt.toISOString().slice(0, 10),                    // YYYY-MM-DD
        dt.toISOString().slice(11, 19),                   // HH:MM:SS (UTC)
        csvCell(t.name),
        csvCell(t.email),
        ((t.amount || 0) / 100).toFixed(2),               // major unit
        (t.currency || "EUR").toUpperCase(),
        csvCell(t.status),
        csvCell(t.payment_status || "paid"),
        csvCell(t.payment_intent_id || ""),
        csvCell(t.stripe_session_id),
        csvCell(t.used_at || ""),
      ].join(",");
    });

    const csv      = [HEADERS.join(","), ...rows].join("\r\n");
    const filename = `rotunda-tickets-${new Date().toISOString().slice(0, 10)}.csv`;

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send("\uFEFF" + csv); // UTF-8 BOM — Excel opens it correctly without mangling

  } catch (err) {
    console.error("❌ /admin/export.csv error:", err);
    res.status(500).json({ error: "Server error" });
  }
});

// ─────────────────────────────────────────────────────────────
// UTILITIES
// ─────────────────────────────────────────────────────────────
function escapeHtml(str = "") {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Wrap a CSV cell value — quotes if it contains comma, quote or newline
function csvCell(val = "") {
  const s = String(val);
  if (s.includes(",") || s.includes('"') || s.includes("\n")) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

// Convert a YYYY-MM-DD date string to end-of-day UTC for inclusive range queries
function toEndOfDay(dateStr = "") {
  return dateStr.length === 10 ? `${dateStr}T23:59:59.999Z` : dateStr;
}

// ─────────────────────────────────────────────────────────────
// START
// ─────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
  console.log(`🌐 Base URL: ${BASE_URL}`);
  console.log(`🔐 Admin routes protected by ADMIN_TOKEN`);
});
