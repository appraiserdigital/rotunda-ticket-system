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
  const net_cents = Math.round(gross_cents * 100 / 118);
  const vat_cents = gross_cents - net_cents;
  return {
    gross_cents, net_cents, vat_cents,
    gross_major: (gross_cents / 100).toFixed(2),
    net_major:   (net_cents   / 100).toFixed(2),
    vat_major:   (vat_cents   / 100).toFixed(2),
  };
}

async function sendTicketEmail(ticket, qrDataUrl) {
  if (!ticket.email) return;
  const ticketUrl = `${BASE_URL}/ticket/${ticket.id}`;
  const expiry    = new Date(ticket.expires_at).toLocaleDateString("en-MT", { day:"2-digit", month:"long", year:"numeric" });
  const amount    = (ticket.amount/100).toLocaleString("en-MT",{ style:"currency", currency:(ticket.currency||"EUR").toUpperCase() });
  try {
    await resend.emails.send({
      from: process.env.EMAIL_FROM,
      to:   ticket.email,
      subject: `Your Ticket — ${ticket.tour_type} | ${FOUNDATION_NAME}`,
      html: `<!DOCTYPE html><html><head><meta charset="UTF-8"/></head>
<body style="margin:0;padding:0;background:#f4f6fa;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:30px 0;">
<tr><td align="center"><table width="560" cellpadding="0" cellspacing="0"
  style="background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08);">
<tr><td style="background:#1B3A6B;padding:28px 32px;">
  <p style="margin:0;color:#B8962E;font-size:12px;letter-spacing:2px;font-weight:bold;">YOUR TICKET</p>
  <p style="margin:8px 0 0;color:#fff;font-size:22px;font-weight:bold;">${FOUNDATION_NAME}</p>
</td></tr>
<tr><td style="padding:28px 32px 0;">
  <p style="margin:0 0 4px;color:#718096;font-size:12px;text-transform:uppercase;">Tour</p>
  <p style="margin:0 0 20px;color:#1B3A6B;font-size:20px;font-weight:bold;">${escapeHtml(ticket.tour_type)}</p>
  <p style="margin:0 0 4px;color:#718096;font-size:12px;text-transform:uppercase;">Name</p>
  <p style="margin:0 0 20px;color:#1A202C;font-size:16px;">${escapeHtml(ticket.name)}</p>
  <table width="100%"><tr>
    <td width="50%"><p style="margin:0 0 4px;color:#718096;font-size:12px;text-transform:uppercase;">Amount Paid</p>
      <p style="margin:0;color:#1A202C;font-size:16px;font-weight:bold;">${amount}</p></td>
    <td width="50%"><p style="margin:0 0 4px;color:#718096;font-size:12px;text-transform:uppercase;">Valid Until</p>
      <p style="margin:0;color:#1A202C;font-size:16px;">${expiry}</p></td>
  </tr></table>
</td></tr>
<tr><td style="padding:28px 32px;text-align:center;">
  <p style="margin:0 0 16px;color:#4A5568;font-size:14px;">Present this QR code at the entrance</p>
  <img src="${qrDataUrl}" width="200" height="200" alt="QR Code"
    style="border:12px solid #fff;border-radius:8px;box-shadow:0 2px 12px rgba(0,0,0,.12);"/>
</td></tr>
<tr><td style="padding:0 32px 28px;text-align:center;">
  <a href="${ticketUrl}" style="display:inline-block;background:#1B3A6B;color:#fff;text-decoration:none;padding:14px 32px;border-radius:8px;font-size:15px;font-weight:bold;">Open My Ticket</a>
</td></tr>
<tr><td style="background:#F4F6FA;padding:20px 32px;border-top:1px solid #E2E8F0;">
  <p style="margin:0;color:#718096;font-size:12px;line-height:1.6;">
    This ticket is valid for one entry and expires on ${expiry}.
    Each QR code is unique and can only be scanned once.
    Keep this email safe — it is your proof of purchase.
  </p>
</td></tr>
</table></td></tr></table></body></html>`,
    });
    console.log(`📧 Ticket email sent to ${ticket.email}`);
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

// SUCCESS
app.get("/success", async (req, res) => {
  const sessionId = req.query.session_id;
  if (!sessionId) return res.status(400).send("Missing session_id.");
  console.log(`🔍 /success — session: ${sessionId}`);
  try {
    const { data: existing, error: lookupErr } = await supabase
      .from("tickets").select("id").eq("stripe_session_id", sessionId).maybeSingle();
    if (lookupErr) { console.error("❌ Supabase lookup:", lookupErr); return res.status(500).send("Database error during lookup."); }
    if (existing) { console.log(`ℹ️  Ticket exists: ${existing.id}`); return res.redirect(`/ticket/${existing.id}`); }

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

    const lineItem = session.line_items?.data?.[0];
    const tourType = lineItem?.price?.product?.name || lineItem?.description || "General Admission";

    const now = new Date();
    const expiresAt = new Date(now);
    expiresAt.setMonth(expiresAt.getMonth() + TICKET_VALIDITY_MONTHS);

    const ticketId = uuidv4();
    const record = {
      id: ticketId,
      stripe_session_id: session.id,
      payment_intent_id: session.payment_intent || null,
      name:     session.customer_details?.name  || "Guest",
      email:    session.customer_details?.email || "",
      tour_type: tourType,
      amount:   session.amount_total,
      currency: (session.currency || "eur").toLowerCase(),
      payment_status: session.payment_status,
      status:   "VALID",
      created_at: now.toISOString(),
      expires_at: expiresAt.toISOString(),
      used_at: null,
    };

    const { error: insertErr } = await supabase.from("tickets").insert([record]);
    if (insertErr) {
      if (insertErr.code === "23505") {
        const { data: race } = await supabase.from("tickets").select("id").eq("stripe_session_id", sessionId).single();
        if (race) return res.redirect(`/ticket/${race.id}`);
      }
      console.error("❌ Supabase insert:", insertErr);
      return res.status(500).send("Database error creating ticket.");
    }

    console.log(`✅ Ticket: ${ticketId} | ${record.email} | ${tourType} | ${record.currency.toUpperCase()} ${(record.amount/100).toFixed(2)}`);
    const checkUrl  = `${BASE_URL}/check/${ticketId}`;
    const qrDataUrl = await QRCode.toDataURL(checkUrl, { errorCorrectionLevel:"H", margin:2, width:300 });
    sendTicketEmail(record, qrDataUrl);
    return res.redirect(`/ticket/${ticketId}`);
  } catch (err) {
    console.error("❌ /success error:", err);
    return res.status(500).send("Unexpected server error.");
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

    const HEADERS=["ticket_id","date","time_utc","name","email","tour_type","gross_eur","net_eur_excl_vat","vat_18pct","currency","ticket_status","payment_status","payment_intent_id","stripe_session_id","expires_at","used_at"];
    const rows=tickets.map(t=>{
      const dt=new Date(t.created_at);
      const v=vatBreakdown(t.amount||0);
      return [csvCell(t.id),dt.toISOString().slice(0,10),dt.toISOString().slice(11,19),csvCell(t.name),csvCell(t.email),csvCell(t.tour_type||"General Admission"),v.gross_major,v.net_major,v.vat_major,(t.currency||"EUR").toUpperCase(),csvCell(t.status),csvCell(t.payment_status||"paid"),csvCell(t.payment_intent_id||""),csvCell(t.stripe_session_id),csvCell(t.expires_at||""),csvCell(t.used_at||"")].join(",");
    });

    const totalGross=tickets.reduce((s,t)=>s+(t.amount||0),0);
    const totals=vatBreakdown(totalGross);
    rows.push("");
    rows.push(`TOTALS,,,,,,${totals.gross_major},${totals.net_major},${totals.vat_major},,,,,,`);
    rows.push(`"VAT Rate: 18% included in all prices. Net = Gross x 100/118. VAT = Gross - Net.",,,,,,,,,,,,,,`);

    const csv=[HEADERS.join(","),...rows].join("\r\n");
    const filename=`rotunda-tickets-${new Date().toISOString().slice(0,10)}.csv`;
    res.setHeader("Content-Type","text/csv; charset=utf-8");
    res.setHeader("Content-Disposition",`attachment; filename="${filename}"`);
    res.send("\uFEFF"+csv);
  } catch(err){ res.status(500).json({error:"Server error"}); }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
  console.log(`🌐 Base URL: ${BASE_URL}`);
});

// ADMIN XLSX EXPORT (formatted)
// GET /admin/export.xlsx?token=xxx&from=...&to=...
const XLSX_HEADERS = [
  "ticket_id","date","time_utc","name","email","tour_type",
  "gross_eur","net_eur_excl_vat","vat_18pct","currency",
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
    PAD + "VAT Rate: 18% included in all prices. Net = Gross × 100/118. VAT = Gross − Net." + PAD
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
