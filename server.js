import express  from "express";
import Stripe   from "stripe";
import QRCode   from "qrcode";
import { v4 as uuidv4 }  from "uuid";
import { createClient }  from "@supabase/supabase-js";
import { Resend }        from "resend";
import ExcelJS           from "exceljs";

const REQUIRED_ENV = [
  "SUPABASE_URL","SUPABASE_KEY","STRIPE_SECRET_KEY",
  "BASE_URL","ADMIN_TOKEN","RESEND_API_KEY","EMAIL_FROM","FOUNDATION_NAME",
];
for (const key of REQUIRED_ENV) {
  if (!process.env[key]) { console.error(`❌ STARTUP FAILED — missing env var: ${key}`); process.exit(1); }
}

const BASE_URL        = process.env.BASE_URL.replace(/\/$/, "");
const FOUNDATION_NAME = process.env.FOUNDATION_NAME;
const TICKET_VALIDITY_MONTHS = 15;

const app  = express();
const PORT = process.env.PORT || 3000;

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);
const stripe   = new Stripe(process.env.STRIPE_SECRET_KEY);
const resend   = new Resend(process.env.RESEND_API_KEY);

app.use(express.json());
app.use(express.static("."));

app.get("/health", (_req, res) => res.json({ status: "ok", ts: new Date().toISOString() }));

function requireAdmin(req, res, next) {
  const token = req.query.token || req.headers["x-admin-token"] || "";
  if (!token || token !== process.env.ADMIN_TOKEN) return res.status(401).json({ error: "Unauthorized" });
  next();
}

function vatBreakdown(gross_cents) {
  const net_cents = Math.round(gross_cents * 100 / 105);
  const vat_cents = gross_cents - net_cents;
  return {
    gross_cents, net_cents, vat_cents,
    gross_major: (gross_cents / 100).toFixed(2),
    net_major:   (net_cents   / 100).toFixed(2),
    vat_major:   (vat_cents   / 100).toFixed(2),
  };
}

async function sendMultiTicketEmail(tickets, totalCents, currency) {
  const first = tickets[0];
  if (!first.email) return;

  const buyerName  = first.name.split(" — Ticket")[0];
  const quantity   = tickets.length;
  const totalStr   = (totalCents / 100).toLocaleString("en-MT", { style:"currency", currency:(currency||"EUR").toUpperCase() });
  const expiry     = new Date(first.expires_at).toLocaleDateString("en-MT", { day:"2-digit", month:"long", year:"numeric" });
  const subject    = quantity > 1
    ? `Your ${quantity} Tickets — ${first.tour_type} | ${FOUNDATION_NAME}`
    : `Your Ticket — ${first.tour_type} | ${FOUNDATION_NAME}`;

  const ticketBlocks = tickets.map(t => {
    const ticketUrl  = `${BASE_URL}/ticket/${t.id}`;
    const unitStr    = (t.amount/100).toLocaleString("en-MT",{style:"currency",currency:(t.currency||"EUR").toUpperCase()});
    return `
      <tr><td style="padding:20px 0;border-bottom:1px solid #E2E8F0;">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="160" align="center" style="padding-right:20px;">
              <img src="${t.qrDataUrl}" width="140" height="140" alt="QR Code"
                style="border:8px solid #fff;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.1);display:block;"/>
            </td>
            <td valign="top">
              <p style="margin:0 0 4px;color:#718096;font-size:11px;text-transform:uppercase;letter-spacing:1px;">${escapeHtml(t.tour_type)}</p>
              <p style="margin:0 0 8px;color:#1A202C;font-size:16px;font-weight:bold;">${escapeHtml(t.name)}</p>
              <p style="margin:0 0 12px;color:#4A5568;font-size:13px;">${unitStr} · Valid until ${expiry}</p>
              <a href="${ticketUrl}" style="display:inline-block;background:#1B3A6B;color:#fff;text-decoration:none;padding:8px 20px;border-radius:6px;font-size:13px;font-weight:bold;">Open Ticket</a>
            </td>
          </tr>
        </table>
      </td></tr>`;
  }).join("");

  try {
    await resend.emails.send({
      from:    process.env.EMAIL_FROM,
      to:      first.email,
      subject,
      html: `<!DOCTYPE html><html><head><meta charset="UTF-8"/></head>
<body style="margin:0;padding:0;background:#f4f6fa;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:30px 0;">
<tr><td align="center">
<table width="580" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08);">

  <tr><td style="background:#1B3A6B;padding:28px 32px;">
    <p style="margin:0;color:#B8962E;font-size:12px;letter-spacing:2px;font-weight:bold;">YOUR ${quantity > 1 ? quantity + " TICKETS" : "TICKET"}</p>
    <p style="margin:8px 0 4px;color:#fff;font-size:22px;font-weight:bold;">${FOUNDATION_NAME}</p>
    <p style="margin:0;color:#BEE3F8;font-size:14px;">${escapeHtml(buyerName)} · ${totalStr} total</p>
  </td></tr>

  <tr><td style="padding:24px 32px 0;">
    <table width="100%" cellpadding="0" cellspacing="0">
      ${ticketBlocks}
    </table>
  </td></tr>

  <tr><td style="background:#F4F6FA;padding:20px 32px;border-top:1px solid #E2E8F0;margin-top:24px;">
    <p style="margin:0;color:#718096;font-size:12px;line-height:1.7;">
      Each QR code is unique and valid for <strong>one entry only</strong>. Tickets expire on ${expiry}.
      Keep this email safe — it is your proof of purchase for all ${quantity} ticket${quantity>1?"s":""}.
    </p>
  </td></tr>

</table>
</td></tr></table>
</body></html>`,
    });
    console.log(`📧 Email sent to ${first.email} — ${quantity} ticket(s)`);
  } catch (err) {
    console.error("⚠️  Email send failed:", err.message);
  }
}

function escapeHtml(str = "") {
  return String(str).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function csvCell(val = "") {
  const s = String(val);
  return (s.includes(",")||s.includes('"')||s.includes("\n")) ? `"${s.replace(/"/g,'""')}"` : s;
}
function toEndOfDay(d="") { return d.length===10 ? `${d}T23:59:59.999Z` : d; }

// SUCCESS — handles single and multi-ticket orders
app.get("/success", async (req, res) => {
  const sessionId = req.query.session_id;
  if (!sessionId) return res.status(400).send("Missing session_id.");
  console.log(`🔍 /success — session: ${sessionId}`);

  try {
    // ── Idempotency: check if any tickets already exist for this session ──
    const { data: existing, error: lookupErr } = await supabase
      .from("tickets").select("id").eq("stripe_session_id", sessionId).limit(1).maybeSingle();
    if (lookupErr) { console.error("❌ Supabase lookup:", lookupErr); return res.status(500).send("Database error during lookup."); }
    if (existing) {
      console.log(`ℹ️  Tickets already exist for session: ${sessionId}`);
      return res.redirect(`/order/${sessionId}`);
    }

    // ── Retrieve session from Stripe ──
    let session;
    try {
      session = await stripe.checkout.sessions.retrieve(sessionId, {
        expand: ["line_items","line_items.data.price.product"],
      });
    } catch (stripeErr) {
      console.error("❌ Stripe retrieve:", stripeErr.message);
      return res.status(502).send(`Stripe error: ${stripeErr.message}`);
    }

    if (session.payment_status !== "paid") return res.status(402).send(`Payment not completed (${session.payment_status}).`);

    // ── Get tour type and quantity ──
    const lineItem  = session.line_items?.data?.[0];
    const tourType  = lineItem?.price?.product?.name || lineItem?.description || "General Admission";
    const quantity  = lineItem?.quantity || 1;
    const unitPrice = quantity > 0 ? Math.round(session.amount_total / quantity) : session.amount_total;

    const buyerName = session.customer_details?.name  || "Guest";
    const email     = session.customer_details?.email || "";
    const currency  = (session.currency || "eur").toLowerCase();
    const now       = new Date();
    const expiresAt = new Date(now);
    expiresAt.setMonth(expiresAt.getMonth() + TICKET_VALIDITY_MONTHS);

    // ── Create one record per ticket ──
    const records = [];
    for (let i = 0; i < quantity; i++) {
      const ticketLabel = quantity > 1 ? `${buyerName} — Ticket ${i + 1} of ${quantity}` : buyerName;
      records.push({
        id:                uuidv4(),
        stripe_session_id: session.id,
        payment_intent_id: session.payment_intent || null,
        name:              ticketLabel,
        email,
        tour_type:         tourType,
        amount:            unitPrice,            // per-ticket price for financial accuracy
        currency,
        payment_status:    session.payment_status,
        status:            "VALID",
        created_at:        now.toISOString(),
        expires_at:        expiresAt.toISOString(),
        used_at:           null,
        order_quantity:    quantity,             // stored for reference
        order_index:       i + 1,               // position in order (1-based)
      });
    }

    const { error: insertErr } = await supabase.from("tickets").insert(records);
    if (insertErr) {
      if (insertErr.code === "23505") {
        console.warn("⚠️  Race condition — redirecting to order page");
        return res.redirect(`/order/${sessionId}`);
      }
      console.error("❌ Supabase insert:", insertErr);
      return res.status(500).send("Database error creating tickets.");
    }

    console.log(`✅ ${quantity} ticket(s) created for ${email} | ${tourType} | ${currency.toUpperCase()} ${(session.amount_total/100).toFixed(2)} total`);

    // ── Generate QR codes and send single email with all tickets ──
    const ticketsWithQR = await Promise.all(records.map(async t => ({
      ...t,
      qrDataUrl: await QRCode.toDataURL(`${BASE_URL}/check/${t.id}`, { errorCorrectionLevel:"H", margin:2, width:220 }),
    })));

    sendMultiTicketEmail(ticketsWithQR, session.amount_total, currency);

    // ── Redirect: single ticket → ticket page, multiple → order page ──
    if (quantity === 1) {
      return res.redirect(`/ticket/${records[0].id}`);
    }
    return res.redirect(`/order/${sessionId}`);

  } catch (err) {
    console.error("❌ /success error:", err);
    return res.status(500).send("Unexpected server error.");
  }
});

// ORDER PAGE — shows all tickets for a multi-ticket purchase
app.get("/order/:sessionId", async (req, res) => {
  try {
    const { data: tickets, error } = await supabase
      .from("tickets")
      .select("*")
      .eq("stripe_session_id", req.params.sessionId)
      .order("order_index", { ascending: true });

    if (error || !tickets?.length) return res.status(404).send("Order not found.");

    const first     = tickets[0];
    const buyerName = first.name.split(" — Ticket")[0];
    const total     = tickets.reduce((s, t) => s + (t.amount || 0), 0);
    const totalStr  = (total / 100).toLocaleString("en-MT", { style:"currency", currency:(first.currency||"EUR").toUpperCase() });
    const dateStr   = new Date(first.created_at).toLocaleDateString("en-MT", { day:"2-digit", month:"long", year:"numeric" });

    const ticketCards = await Promise.all(tickets.map(async t => {
      const qr = await QRCode.toDataURL(`${BASE_URL}/check/${t.id}`, { errorCorrectionLevel:"H", margin:2, width:180 });
      const isExpired = t.expires_at && new Date(t.expires_at) < new Date();
      const statusCol = (t.status==="VALID" && !isExpired) ? "#22c55e" : "#ef4444";
      const label     = isExpired ? "EXPIRED" : t.status;
      return `
        <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:16px;padding:24px;text-align:center;flex:1;min-width:200px;max-width:240px;">
          <p style="font-size:13px;color:#93c5fd;font-weight:bold;margin-bottom:12px;">${escapeHtml(t.tour_type)}</p>
          <div style="background:#fff;border-radius:10px;padding:10px;display:inline-block;margin-bottom:14px;">
            <img src="${qr}" width="160" height="160" alt="QR"/>
          </div>
          <p style="font-size:14px;color:#fff;font-weight:bold;margin-bottom:4px;">${escapeHtml(t.name)}</p>
          <p style="font-size:12px;color:#999;margin-bottom:10px;">${(t.amount/100).toLocaleString("en-MT",{style:"currency",currency:(t.currency||"EUR").toUpperCase()})}</p>
          <span style="display:inline-block;padding:4px 16px;border-radius:12px;font-size:11px;font-weight:bold;color:${statusCol};border:1px solid ${statusCol};background:${statusCol}22;" id="badge-${t.id}">${label}</span>
          <br><a href="/ticket/${t.id}" style="display:inline-block;margin-top:12px;font-size:12px;color:#3b82f6;text-decoration:none;">Open full ticket →</a>
        </div>`;
    }));

    return res.send(`<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Your Order — ${escapeHtml(FOUNDATION_NAME)}</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{font-family:Arial,sans-serif;background:#0f0f1a;color:#e5e5e5;min-height:100vh;padding:30px 20px;}
.wrap{max-width:900px;margin:0 auto;}
.header{text-align:center;margin-bottom:32px;}
.header h1{font-size:24px;color:#fff;margin-bottom:6px;}
.header p{font-size:14px;color:#888;}
.cards{display:flex;flex-wrap:wrap;gap:16px;justify-content:center;}
</style></head><body>
<div class="wrap">
  <div class="header">
    <p style="color:#B8962E;font-size:12px;letter-spacing:2px;font-weight:bold;text-transform:uppercase;margin-bottom:8px;">Your Order</p>
    <h1>🎟 ${escapeHtml(FOUNDATION_NAME)}</h1>
    <p>${escapeHtml(buyerName)} · ${dateStr} · ${tickets.length} ticket${tickets.length>1?"s":""} · ${totalStr}</p>
  </div>
  <div class="cards">${ticketCards.join("")}</div>
  <p style="text-align:center;margin-top:28px;font-size:12px;color:#444;">Each QR code is unique and valid for one entry. Bookmark this page to access your tickets again.</p>
</div>
<script>
// Poll each ticket every 4 seconds and update its badge
const ticketIds = ${JSON.stringify(tickets.map(t => t.id))};
const BASE = "${BASE_URL}";

function applyBadge(id, status) {
  const badge = document.getElementById("badge-" + id);
  if (!badge) return;
  if (status === "USED" || status === "ALREADY_USED") {
    badge.style.cssText = "display:inline-block;padding:4px 16px;border-radius:12px;font-size:11px;font-weight:bold;color:#ef4444;border:1px solid #ef4444;background:#ef444422;";
    badge.textContent = "USED";
  } else if (status === "EXPIRED") {
    badge.style.cssText = "display:inline-block;padding:4px 16px;border-radius:12px;font-size:11px;font-weight:bold;color:#f59e0b;border:1px solid #f59e0b;background:#f59e0b22;";
    badge.textContent = "EXPIRED";
  }
}

const intervals = {};
ticketIds.forEach(id => {
  let last = "VALID";
  intervals[id] = setInterval(() => {
    fetch(BASE + "/check/" + id)
      .then(r => r.json())
      .then(d => {
        if (d.status !== last) {
          last = d.status;
          applyBadge(id, d.status);
          // Stop polling once terminal state reached
          if (d.status === "USED" || d.status === "ALREADY_USED" || d.status === "EXPIRED") {
            clearInterval(intervals[id]);
          }
        }
      }).catch(() => {});
  }, 4000);
});
</script>
</body></html>`);

  } catch (err) {
    console.error("❌ /order error:", err);
    return res.status(500).send("Error loading order.");
  }
});

// TICKET PAGE
app.get("/ticket/:id", async (req, res) => {
  try {
    const { data: ticket, error } = await supabase.from("tickets").select("*").eq("id", req.params.id).single();
    if (error || !ticket) return res.status(404).send("Ticket not found.");

    const checkUrl  = `${BASE_URL}/check/${ticket.id}`;
    const qrDataUrl = await QRCode.toDataURL(checkUrl, { errorCorrectionLevel:"H", margin:2, width:300 });

    const isExpired    = ticket.expires_at && new Date(ticket.expires_at) < new Date();
    const statusLabel  = isExpired ? "EXPIRED" : ticket.status;
    const statusColour = (ticket.status==="VALID" && !isExpired) ? "#22c55e" : isExpired ? "#f59e0b" : "#ef4444";
    const amount = ticket.amount!=null ? (ticket.amount/100).toLocaleString("en-MT",{style:"currency",currency:(ticket.currency||"EUR").toUpperCase()}) : "";
    const dateStr   = new Date(ticket.created_at).toLocaleDateString("en-MT",{day:"2-digit",month:"long",year:"numeric"});
    const expiryStr = ticket.expires_at ? new Date(ticket.expires_at).toLocaleDateString("en-MT",{day:"2-digit",month:"long",year:"numeric"}) : "";

    return res.send(`<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Your Ticket — ${escapeHtml(FOUNDATION_NAME)}</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{font-family:Arial,sans-serif;background:#0f0f1a;color:#e5e5e5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;}
.ticket{background:#1a1a2e;border:1px solid #2a2a4a;border-radius:20px;padding:32px 24px;max-width:380px;width:100%;text-align:center;box-shadow:0 25px 60px rgba(0,0,0,.6);}
.title{font-size:22px;font-weight:bold;color:#fff;margin-bottom:4px;}
.venue{font-size:13px;color:#888;margin-bottom:8px;letter-spacing:1px;text-transform:uppercase;}
.tour{font-size:15px;color:#93c5fd;font-weight:bold;margin-bottom:20px;}
.qr-wrap{background:#fff;border-radius:12px;padding:14px;display:inline-block;margin-bottom:22px;}
.qr-wrap img{display:block;width:200px;height:200px;}
.name{font-size:20px;font-weight:bold;color:#fff;margin-bottom:6px;}
.detail{font-size:13px;color:#999;margin-bottom:3px;}
.expiry{font-size:12px;color:#64748b;margin-top:6px;}
.status{display:inline-block;margin-top:16px;padding:6px 22px;border-radius:20px;font-size:13px;font-weight:bold;letter-spacing:1px;color:${statusColour};border:1px solid ${statusColour};background:${statusColour}22;}
.footer{margin-top:20px;font-size:11px;color:#444;}
</style></head><body>
<div class="ticket">
<p class="title">🎟 Ticket</p>
<p class="venue">${escapeHtml(FOUNDATION_NAME)}</p>
<p class="tour">${escapeHtml(ticket.tour_type||"General Admission")}</p>
<div class="qr-wrap"><img src="${qrDataUrl}" alt="Ticket QR Code"/></div>
<p class="name">${escapeHtml(ticket.name)}</p>
<p class="detail">${escapeHtml(ticket.email)}</p>
<p class="detail">${dateStr}${amount?" · "+amount:""}</p>
${expiryStr?`<p class="expiry">Valid until ${expiryStr}</p>`:""}
<span class="status" id="status-badge">${statusLabel}</span>
<p class="footer" id="footer-text">Present this QR code at the entrance</p>
</div>
<script>
const checkUrl="${BASE_URL}/check/${ticket.id}";
let lastStatus="${ticket.status}";
const badge=document.getElementById("status-badge");
const footer=document.getElementById("footer-text");
function applyStatus(s){
  if(s==="VALID"){badge.style.cssText="color:#22c55e;border-color:#22c55e;background:#22c55e22;";badge.textContent="VALID";footer.textContent="Present this QR code at the entrance";}
  else if(s==="USED"||s==="ALREADY_USED"){badge.style.cssText="color:#ef4444;border-color:#ef4444;background:#ef444422;";badge.textContent="USED";footer.textContent="This ticket has already been scanned at the entrance";clearInterval(poll);}
  else if(s==="EXPIRED"){badge.style.cssText="color:#f59e0b;border-color:#f59e0b;background:#f59e0b22;";badge.textContent="EXPIRED";footer.textContent="This ticket has expired";clearInterval(poll);}
}
const poll=setInterval(()=>{fetch(checkUrl).then(r=>r.json()).then(d=>{if(d.status!==lastStatus){lastStatus=d.status;applyStatus(d.status);}}).catch(()=>{});},4000);
</script></body></html>`);
  } catch (err) {
    console.error("❌ /ticket error:", err);
    return res.status(500).send("Error loading ticket.");
  }
});

// CHECK
app.get("/check/:id", async (req, res) => {
  try {
    const { data: ticket, error } = await supabase.from("tickets").select("id,status,name,expires_at,tour_type").eq("id", req.params.id).single();
    if (error||!ticket) return res.json({status:"INVALID"});
    if (ticket.status==="USED") return res.json({status:"ALREADY_USED",name:ticket.name,tour:ticket.tour_type});
    if (ticket.expires_at&&new Date(ticket.expires_at)<new Date()) return res.json({status:"EXPIRED",name:ticket.name,tour:ticket.tour_type});
    return res.json({status:"VALID",name:ticket.name,tour:ticket.tour_type});
  } catch(err){ return res.json({status:"ERROR"}); }
});

// USE
app.get("/use/:id", async (req, res) => {
  try {
    const { data: ticket, error } = await supabase.from("tickets").select("id,status,name,expires_at,tour_type").eq("id", req.params.id).single();
    if (error||!ticket) return res.json({status:"INVALID"});
    if (ticket.status==="USED") return res.json({status:"ALREADY_USED",name:ticket.name,tour:ticket.tour_type});
    if (ticket.expires_at&&new Date(ticket.expires_at)<new Date()) return res.json({status:"EXPIRED",name:ticket.name});
    const { error: updateErr } = await supabase.from("tickets").update({status:"USED",used_at:new Date().toISOString()}).eq("id", req.params.id);
    if (updateErr) { console.error("❌ Update error:",updateErr); return res.json({status:"ERROR"}); }
    console.log(`✅ Used: ${req.params.id} | ${ticket.name} | ${ticket.tour_type}`);
    return res.json({status:"USED",name:ticket.name,tour:ticket.tour_type});
  } catch(err){ return res.json({status:"ERROR"}); }
});

// ADMIN SUMMARY
app.get("/admin/summary", requireAdmin, async (req, res) => {
  try {
    let q = supabase.from("tickets").select("amount,currency,status,tour_type,created_at");
    if (req.query.from) q = q.gte("created_at", req.query.from);
    if (req.query.to)   q = q.lte("created_at", toEndOfDay(req.query.to));
    const { data: tickets, error } = await q;
    if (error) return res.status(500).json({error:error.message});

    const total = tickets.length;
    const used  = tickets.filter(t=>t.status==="USED").length;
    const valid = tickets.filter(t=>t.status==="VALID").length;

    const byCurrency = {};
    for (const t of tickets) {
      const cur = (t.currency||"eur").toUpperCase();
      if (!byCurrency[cur]) byCurrency[cur]={count:0,gross_cents:0};
      byCurrency[cur].count++;
      byCurrency[cur].gross_cents += t.amount||0;
    }
    for (const cur of Object.keys(byCurrency)) {
      const v = vatBreakdown(byCurrency[cur].gross_cents);
      Object.assign(byCurrency[cur], {gross_major:v.gross_major,net_major:v.net_major,vat_major:v.vat_major});
    }

    const byTour = {};
    for (const t of tickets) {
      const tour = t.tour_type||"General Admission";
      if (!byTour[tour]) byTour[tour]={count:0,used:0,gross_cents:0,currency:t.currency||"eur"};
      byTour[tour].count++;
      byTour[tour].gross_cents += t.amount||0;
      if (t.status==="USED") byTour[tour].used++;
    }
    for (const tour of Object.keys(byTour)) {
      const v = vatBreakdown(byTour[tour].gross_cents);
      Object.assign(byTour[tour], {gross_major:v.gross_major,net_major:v.net_major,vat_major:v.vat_major});
    }

    res.json({total_tickets:total,used_tickets:used,valid_tickets:valid,by_currency:byCurrency,by_tour:byTour,generated_at:new Date().toISOString()});
  } catch(err){ res.status(500).json({error:"Server error"}); }
});

// ADMIN DAILY
app.get("/admin/daily", requireAdmin, async (req, res) => {
  try {
    let q = supabase.from("tickets").select("amount,currency,status,tour_type,created_at").order("created_at",{ascending:true});
    if (req.query.from) q = q.gte("created_at", req.query.from);
    if (req.query.to)   q = q.lte("created_at", toEndOfDay(req.query.to));
    const { data: tickets, error } = await q;
    if (error) return res.status(500).json({error:error.message});
    const byDate={};
    for (const t of tickets) {
      const date=t.created_at.slice(0,10);
      if (!byDate[date]) byDate[date]={date,count:0,gross_cents:0,currency:(t.currency||"eur").toUpperCase()};
      byDate[date].count++;
      byDate[date].gross_cents+=t.amount||0;
    }
    const rows=Object.values(byDate).map(d=>{const v=vatBreakdown(d.gross_cents);return{...d,gross_major:v.gross_major,net_major:v.net_major,vat_major:v.vat_major};});
    res.json(rows);
  } catch(err){ res.status(500).json({error:"Server error"}); }
});

// ADMIN TICKETS
app.get("/admin/tickets", requireAdmin, async (req, res) => {
  try {
    let q = supabase.from("tickets").select("*").order("created_at",{ascending:false}).limit(1000);
    if (req.query.status)    q = q.eq("status", req.query.status);
    if (req.query.tour_type) q = q.eq("tour_type", req.query.tour_type);
    if (req.query.from)      q = q.gte("created_at", req.query.from);
    if (req.query.to)        q = q.lte("created_at", toEndOfDay(req.query.to));
    if (req.query.email)     q = q.ilike("email", `%${req.query.email}%`);
    const { data, error } = await q;
    if (error) return res.status(500).json({error:error.message});
    res.json(data);
  } catch(err){ res.status(500).json({error:"Server error"}); }
});

// ADMIN CSV EXPORT
app.get("/admin/export.csv", requireAdmin, async (req, res) => {
  try {
    let q = supabase.from("tickets").select("*").order("created_at",{ascending:true});
    if (req.query.status)    q = q.eq("status", req.query.status);
    if (req.query.tour_type) q = q.eq("tour_type", req.query.tour_type);
    if (req.query.from)      q = q.gte("created_at", req.query.from);
    if (req.query.to)        q = q.lte("created_at", toEndOfDay(req.query.to));
    const { data: tickets, error } = await q;
    if (error) return res.status(500).json({error:error.message});

    const HEADERS=["ticket_id","date","time_utc","name","email","tour_type","gross_eur","net_eur_excl_vat","vat_5pct","currency","ticket_status","payment_status","payment_intent_id","stripe_session_id","expires_at","used_at"];
    const rows=tickets.map(t=>{
      const dt=new Date(t.created_at);
      const v=vatBreakdown(t.amount||0);
      return [csvCell(t.id),dt.toISOString().slice(0,10),dt.toISOString().slice(11,19),csvCell(t.name),csvCell(t.email),csvCell(t.tour_type||"General Admission"),v.gross_major,v.net_major,v.vat_major,(t.currency||"EUR").toUpperCase(),csvCell(t.status),csvCell(t.payment_status||"paid"),csvCell(t.payment_intent_id||""),csvCell(t.stripe_session_id),csvCell(t.expires_at||""),csvCell(t.used_at||"")].join(",");
    });

    const totalGross=tickets.reduce((s,t)=>s+(t.amount||0),0);
    const totals=vatBreakdown(totalGross);
    rows.push("");
    rows.push(`TOTALS,,,,,,${totals.gross_major},${totals.net_major},${totals.vat_major},,,,,,`);
    rows.push(`"VAT Rate: 5% included in all prices (Malta VAT rate for tours). Net = Gross x 100/105. VAT = Gross - Net.",,,,,,,,,,,,,,`);

    const csv=[HEADERS.join(","),...rows].join("\r\n");
    const filename=`rotunda-tickets-${new Date().toISOString().slice(0,10)}.csv`;
    res.setHeader("Content-Type","text/csv; charset=utf-8");
    res.setHeader("Content-Disposition",`attachment; filename="${filename}"`);
    res.send("\uFEFF"+csv);
  } catch(err){ res.status(500).json({error:"Server error"}); }
});

// ADMIN XLSX EXPORT (formatted)
// GET /admin/export.xlsx?token=xxx&from=...&to=...
const XLSX_HEADERS = [
  "ticket_id","date","time_utc","name","email","tour_type",
  "gross_eur","net_eur_excl_vat","vat_5pct","currency",
  "ticket_status","payment_status","payment_intent_id","stripe_session_id",
  "expires_at","used_at"
];
// Column widths in chars — stripe_session_id (col 14) is fixed wide, not auto
const XLSX_WIDTHS = [38,14,12,26,36,28,13,18,12,10,14,16,36,52,22,22];
const PAD = "   "; // 3 spaces before and after each value

async function buildXlsx(tickets) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Rotunda Ticket System";
  const ws = wb.addWorksheet("Tickets", { views: [{ state:"frozen", ySplit:1 }] });

  ws.columns = XLSX_HEADERS.map((h, i) => ({
    key:   h,
    width: XLSX_WIDTHS[i],
  }));

  // ── Header row ───────────────────────────────────────────────
  const headerRow = ws.addRow(
    XLSX_HEADERS.map(h => PAD + h.replace(/_/g, " ").toUpperCase() + PAD)
  );
  headerRow.height = 22;
  headerRow.eachCell(cell => {
    cell.font      = { bold:true, name:"Arial", size:10, color:{ argb:"FF1B3A6B" } };
    cell.fill      = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFD6E4F0" } };
    cell.alignment = { vertical:"middle", horizontal:"left" };
    cell.border    = { bottom:{ style:"medium", color:{ argb:"FF1B3A6B" } } };
  });

  // ── Data rows ────────────────────────────────────────────────
  const LIGHT_BLUE = "FFE8F4FB";
  const WHITE      = "FFFFFFFF";

  tickets.forEach((t, i) => {
    const dt  = new Date(t.created_at);
    const v   = vatBreakdown(t.amount || 0);
    const row = ws.addRow([
      PAD + (t.id                          || "") + PAD,
      PAD + dt.toISOString().slice(0,10)          + PAD,
      PAD + dt.toISOString().slice(11,19)         + PAD,
      PAD + (t.name                        || "") + PAD,
      PAD + (t.email                       || "") + PAD,
      PAD + (t.tour_type || "General Admission")  + PAD,
      PAD + v.gross_major                         + PAD,
      PAD + v.net_major                           + PAD,
      PAD + v.vat_major                           + PAD,
      PAD + (t.currency  || "EUR").toUpperCase()  + PAD,
      PAD + (t.status                      || "") + PAD,
      PAD + (t.payment_status || "paid")          + PAD,
      PAD + (t.payment_intent_id           || "") + PAD,
      PAD + (t.stripe_session_id           || "") + PAD,
      PAD + (t.expires_at                  || "") + PAD,
      PAD + (t.used_at                     || "") + PAD,
    ]);

    const fill = i % 2 === 0 ? LIGHT_BLUE : WHITE;
    row.height = 18;
    row.eachCell(cell => {
      cell.fill      = { type:"pattern", pattern:"solid", fgColor:{ argb:fill } };
      cell.font      = { name:"Arial", size:10 };
      cell.alignment = { vertical:"middle", horizontal:"left" };
    });
  });

  // ── Blank row ────────────────────────────────────────────────
  ws.addRow([]);

  // ── Totals row ───────────────────────────────────────────────
  const totalGross = tickets.reduce((s,t) => s + (t.amount||0), 0);
  const totals     = vatBreakdown(totalGross);
  const dataStart  = 2;
  const dataEnd    = tickets.length + 1;

  const totalsRow = ws.addRow([
    PAD + "TOTALS" + PAD, "", "", "", "", "",
    PAD + totals.gross_major + PAD,
    PAD + totals.net_major   + PAD,
    PAD + totals.vat_major   + PAD,
    "", "", "", "", "", "", ""
  ]);
  totalsRow.height = 20;
  totalsRow.eachCell((cell, colIdx) => {
    cell.font   = { bold:true, name:"Arial", size:10, color:{ argb:"FF1B3A6B" } };
    cell.fill   = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFDBEAFE" } };
    cell.border = { top:{ style:"medium", color:{ argb:"FF1B3A6B" } } };
    cell.alignment = { vertical:"middle", horizontal:"left" };
  });

  // ── VAT note ─────────────────────────────────────────────────
  ws.addRow([]);
  const noteRow = ws.addRow([
    PAD + "VAT Rate: 5% included in all prices (Malta VAT rate for tours). Net = Gross × 100/105. VAT = Gross − Net." + PAD
  ]);
  noteRow.getCell(1).font      = { italic:true, name:"Arial", size:9, color:{ argb:"FF718096" } };
  noteRow.getCell(1).alignment = { vertical:"middle", horizontal:"left" };

  return wb;
}

app.get("/admin/export.xlsx", requireAdmin, async (req, res) => {
  try {
    let q = supabase.from("tickets").select("*").order("created_at",{ascending:true});
    if (req.query.status)    q = q.eq("status",    req.query.status);
    if (req.query.tour_type) q = q.eq("tour_type", req.query.tour_type);
    if (req.query.from)      q = q.gte("created_at", req.query.from);
    if (req.query.to)        q = q.lte("created_at", toEndOfDay(req.query.to));

    const { data: tickets, error } = await q;
    if (error) return res.status(500).json({ error: error.message });

    const wb       = await buildXlsx(tickets);
    const filename = `rotunda-tickets-${new Date().toISOString().slice(0,10)}.xlsx`;

    res.setHeader("Content-Type",        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

    await wb.xlsx.write(res);
    res.end();

  } catch(err) {
    console.error("❌ /admin/export.xlsx error:", err);
    res.status(500).json({ error: "Server error" });
  }
});

// DAILY CLOSURE REPORT
// GET /admin/day-report?token=xxx&from=YYYY-MM-DD&to=YYYY-MM-DD
// Returns a printer-friendly HTML page
app.get("/admin/day-report", requireAdmin, async (req, res) => {
  try {
    const from = req.query.from || new Date().toISOString().slice(0, 10);
    const to   = req.query.to   || from;

    let q = supabase
      .from("tickets")
      .select("*")
      .gte("created_at", from)
      .lte("created_at", toEndOfDay(to))
      .order("created_at", { ascending: true });

    if (req.query.tour_type) q = q.eq("tour_type", req.query.tour_type);

    const { data: tickets, error } = await q;
    if (error) return res.status(500).send("Database error");

    const total        = tickets.length;
    const arrived      = tickets.filter(t => t.status === "USED").length;
    const noShow       = tickets.filter(t => t.status === "VALID").length;
    const totalGross   = tickets.reduce((s, t) => s + (t.amount || 0), 0);
    const vat          = vatBreakdown(totalGross);

    const dateLabel = from === to
      ? new Date(from).toLocaleDateString("en-MT", { weekday:"long", day:"2-digit", month:"long", year:"numeric" })
      : `${new Date(from).toLocaleDateString("en-MT", { day:"2-digit", month:"long", year:"numeric" })} — ${new Date(to).toLocaleDateString("en-MT", { day:"2-digit", month:"long", year:"numeric" })}`;

    const generatedAt = new Date().toLocaleString("en-MT", {
      day:"2-digit", month:"long", year:"numeric",
      hour:"2-digit", minute:"2-digit"
    });

    // Group by tour
    const byTour = {};
    for (const t of tickets) {
      const tour = t.tour_type || "General Admission";
      if (!byTour[tour]) byTour[tour] = { count:0, arrived:0, gross:0 };
      byTour[tour].count++;
      byTour[tour].gross += t.amount || 0;
      if (t.status === "USED") byTour[tour].arrived++;
    }

    const tourRows = Object.entries(byTour).map(([tour, d]) => {
      const v = vatBreakdown(d.gross);
      return `<tr>
        <td>${escapeHtml(tour)}</td>
        <td class="num">${d.count}</td>
        <td class="num">${d.arrived}</td>
        <td class="num">${d.count - d.arrived}</td>
        <td class="num">€${v.gross_major}</td>
        <td class="num">€${v.net_major}</td>
        <td class="num">€${v.vat_major}</td>
      </tr>`;
    }).join("");

    const ticketRows = tickets.map((t, i) => {
      const dt      = new Date(t.created_at);
      const timeStr = dt.toLocaleTimeString("en-MT", { hour:"2-digit", minute:"2-digit" });
      const v       = vatBreakdown(t.amount || 0);
      const arrived = t.status === "USED";
      const usedTime = t.used_at
        ? new Date(t.used_at).toLocaleTimeString("en-MT", { hour:"2-digit", minute:"2-digit" })
        : "—";
      return `<tr class="${arrived ? "" : "no-show"}">
        <td class="num">${i + 1}</td>
        <td>${timeStr}</td>
        <td>${escapeHtml(t.name)}</td>
        <td>${escapeHtml(t.email)}</td>
        <td>${escapeHtml(t.tour_type || "General Admission")}</td>
        <td class="num">€${v.gross_major}</td>
        <td class="num ${arrived ? "arrived" : "noshow"}">${arrived ? "✓ ARRIVED " + usedTime : "NO SHOW"}</td>
      </tr>`;
    }).join("");

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <title>Daily Closure Report — ${escapeHtml(FOUNDATION_NAME)}</title>
  <style>
    @media print {
      .no-print { display: none !important; }
      body { margin: 0; }
    }
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; font-size: 12px; color: #1a202c; background: #fff; padding: 30px; }

    .no-print {
      background: #1B3A6B; color: #fff; padding: 12px 20px;
      border-radius: 8px; margin-bottom: 24px; display: flex;
      align-items: center; gap: 16px;
    }
    .no-print button {
      background: #B8962E; color: #fff; border: none; padding: 8px 20px;
      border-radius: 6px; font-size: 13px; font-weight: bold; cursor: pointer;
    }
    .no-print span { font-size: 14px; font-weight: bold; }

    .header { border-bottom: 3px solid #1B3A6B; padding-bottom: 14px; margin-bottom: 20px; }
    .header h1 { font-size: 18px; color: #1B3A6B; margin-bottom: 2px; }
    .header .sub { font-size: 12px; color: #718096; }
    .header .date { font-size: 15px; font-weight: bold; color: #1B3A6B; margin-top: 6px; }

    .notice {
      background: #EBF8FF; border: 1px solid #BEE3F8;
      border-radius: 6px; padding: 10px 14px; margin-bottom: 20px;
      font-size: 11px; color: #2C5282; line-height: 1.6;
    }
    .notice strong { display: block; margin-bottom: 2px; font-size: 12px; }

    .summary-grid {
      display: grid; grid-template-columns: repeat(4, 1fr);
      gap: 12px; margin-bottom: 20px;
    }
    .summary-box {
      border: 1px solid #E2E8F0; border-radius: 6px;
      padding: 10px 12px; text-align: center;
    }
    .summary-box .lbl { font-size: 10px; color: #718096; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 4px; }
    .summary-box .val { font-size: 20px; font-weight: bold; color: #1B3A6B; }
    .summary-box.green .val { color: #166534; }
    .summary-box.orange .val { color: #92400e; }

    .vat-row {
      display: flex; gap: 12px; margin-bottom: 20px;
      background: #F4F6FA; border: 1px solid #E2E8F0;
      border-radius: 6px; padding: 10px 14px;
      font-size: 11px; color: #4A5568;
    }
    .vat-row span { margin-right: 24px; }
    .vat-row strong { color: #1B3A6B; }

    h2 { font-size: 13px; color: #1B3A6B; border-bottom: 1px solid #E2E8F0;
         padding-bottom: 4px; margin-bottom: 10px; margin-top: 20px; }

    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 11px; }
    th { background: #1B3A6B; color: #fff; padding: 7px 8px; text-align: left; font-size: 10px; text-transform: uppercase; letter-spacing: .4px; }
    td { padding: 6px 8px; border-bottom: 1px solid #E2E8F0; }
    tr:nth-child(even) td { background: #F9FAFB; }
    tr.no-show td { color: #9CA3AF; }
    .num { text-align: right; font-variant-numeric: tabular-nums; }
    .arrived { color: #166534; font-weight: bold; }
    .noshow  { color: #9CA3AF; }

    .totals-row td { font-weight: bold; border-top: 2px solid #1B3A6B; background: #EBF8FF !important; }

    .footer {
      margin-top: 30px; border-top: 1px solid #E2E8F0;
      padding-top: 12px; font-size: 10px; color: #9CA3AF;
      display: flex; justify-content: space-between;
    }
    .sign-line {
      margin-top: 30px; display: flex; gap: 40px;
    }
    .sign-box { flex: 1; }
    .sign-box .line { border-bottom: 1px solid #1B3A6B; margin-bottom: 4px; height: 30px; }
    .sign-box .lbl  { font-size: 10px; color: #718096; }
  </style>
</head>
<body>

  <div class="no-print">
    <span>📋 Daily Closure Report — Ready to Print</span>
    <button onclick="window.print()">🖨 Print / Save as PDF</button>
  </div>

  <div class="header">
    <h1>${escapeHtml(FOUNDATION_NAME)}</h1>
    <div class="sub">Online Pre-Booking Closure Report</div>
    <div class="date">${dateLabel}</div>
    <div class="sub" style="margin-top:4px;">Generated: ${generatedAt}</div>
  </div>

  <div class="notice">
    <strong>⚠️ Note for Accountant &amp; Reception</strong>
    All transactions listed below have been collected in full via Stripe online payment at the time of booking.
    These amounts do <strong>not</strong> require processing through the till or POS terminal.
    Fiscal receipts issued at reception for these visitors should be marked <strong>"Pre-Paid Online — Stripe Ref: see attached"</strong>.
    The Stripe payment reference for each ticket is available in the full Excel export.
  </div>

  <div class="summary-grid">
    <div class="summary-box">
      <div class="lbl">Online Bookings</div>
      <div class="val">${total}</div>
    </div>
    <div class="summary-box green">
      <div class="lbl">Arrived</div>
      <div class="val">${arrived}</div>
    </div>
    <div class="summary-box orange">
      <div class="lbl">No Show</div>
      <div class="val">${noShow}</div>
    </div>
    <div class="summary-box">
      <div class="lbl">Total Collected</div>
      <div class="val">€${vat.gross_major}</div>
    </div>
  </div>

  <div class="vat-row">
    <span>Gross Revenue: <strong>€${vat.gross_major}</strong></span>
    <span>Net (excl. 5% VAT): <strong>€${vat.net_major}</strong></span>
    <span>VAT @ 5%: <strong>€${vat.vat_major}</strong></span>
    <span style="margin-left:auto;font-style:italic;">All amounts are VAT-inclusive at the Malta tour rate of 5%</span>
  </div>

  ${Object.keys(byTour).length > 1 ? `
  <h2>Revenue by Tour Type</h2>
  <table>
    <thead><tr>
      <th>Tour</th><th class="num">Bookings</th><th class="num">Arrived</th>
      <th class="num">No Show</th><th class="num">Gross (€)</th>
      <th class="num">Net (€)</th><th class="num">VAT 5% (€)</th>
    </tr></thead>
    <tbody>${tourRows}</tbody>
  </table>` : ""}

  <h2>Visitor List</h2>
  <table>
    <thead><tr>
      <th>#</th><th>Booked At</th><th>Name</th><th>Email</th>
      <th>Tour</th><th class="num">Amount</th><th>Status</th>
    </tr></thead>
    <tbody>
      ${ticketRows}
      <tr class="totals-row">
        <td colspan="5">TOTAL</td>
        <td class="num">€${vat.gross_major}</td>
        <td></td>
      </tr>
    </tbody>
  </table>

  <div class="sign-line">
    <div class="sign-box">
      <div class="line"></div>
      <div class="lbl">Prepared by (Reception)</div>
    </div>
    <div class="sign-box">
      <div class="line"></div>
      <div class="lbl">Date &amp; Time</div>
    </div>
    <div class="sign-box">
      <div class="line"></div>
      <div class="lbl">Verified by (Accounts)</div>
    </div>
  </div>

  <div class="footer">
    <span>${escapeHtml(FOUNDATION_NAME)} — Confidential Financial Document</span>
    <span>Powered by App-Raiser Digital</span>
  </div>

</body>
</html>`;

    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.send(html);

  } catch(err) {
    console.error("❌ /admin/day-report error:", err);
    res.status(500).send("Server error generating report");
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
  console.log(`🌐 Base URL: ${BASE_URL}`);
});
