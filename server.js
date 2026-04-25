import express from 'express';
import Stripe from 'stripe';
import { createClient } from '@supabase/supabase-js';
import QRCode from 'qrcode';

const app = express();
const stripe = new Stripe(process.env.STRIPE_SECRET_KEY);

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
);

// ==========================
// STATIC FILES
// ==========================
app.use(express.static('public'));

// IMPORTANT: JSON AFTER webhook
app.use(express.json());

// ==========================
// STRIPE WEBHOOK
// ==========================
app.post('/webhook', express.raw({ type: 'application/json' }), async (req, res) => {
  const sig = req.headers['stripe-signature'];

  let event;

  try {
    event = stripe.webhooks.constructEvent(
      req.body,
      sig,
      process.env.STRIPE_WEBHOOK_SECRET
    );
  } catch (err) {
    console.error('Webhook error:', err.message);
    return res.status(400).send(`Webhook Error`);
  }

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object;

    try {
      // prevent duplicates
      const { data: existing } = await supabase
        .from('tickets')
        .select('*')
        .eq('stripe_session_id', session.id)
        .maybeSingle();

      if (existing) {
        console.log('Duplicate prevented');
        return res.status(200).send();
      }

      // create ticket
      const { data: ticket, error } = await supabase
        .from('tickets')
        .insert({
          stripe_session_id: session.id,
          status: 'VALID'
        })
        .select()
        .single();

      if (error) {
        console.error('Insert error:', error);
      } else {
        console.log('Ticket created:', ticket.id);
      }

    } catch (err) {
      console.error('Webhook processing error:', err);
    }
  }

  res.status(200).send();
});

// ==========================
// 🔥 THIS WAS MISSING — FIX
// ==========================
app.get('/ticket-by-session/:sessionId', async (req, res) => {
  const { sessionId } = req.params;

  try {
    const { data } = await supabase
      .from('tickets')
      .select('*')
      .eq('stripe_session_id', sessionId)
      .limit(1);

    if (!data || data.length === 0) {
      return res.status(404).json({ status: 'NOT_READY' });
    }

    res.json({
      status: 'READY',
      ticket: data[0]
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ status: 'ERROR' });
  }
});

// ==========================
// TICKET PAGE
// ==========================
app.get('/ticket/:id', async (req, res) => {
  const { id } = req.params;

  const { data: ticket } = await supabase
    .from('tickets')
    .select('*')
    .eq('id', id)
    .maybeSingle();

  if (!ticket) {
    return res.send('<h1>Invalid Ticket</h1>');
  }

  const qr = await QRCode.toDataURL(
    `${process.env.BASE_URL}/check/${ticket.id}`
  );

  res.send(`
    <html>
      <body style="font-family:Arial;text-align:center;padding:40px;">
        <h1>Mosta Rotunda Ticket</h1>
        <p>Status: ${ticket.status}</p>
        <img src="${qr}" />
      </body>
    </html>
  `);
});

// ==========================
// CHECK (SCANNER)
// ==========================
app.get('/check/:id', async (req, res) => {
  const { id } = req.params;

  const { data: ticket } = await supabase
    .from('tickets')
    .select('*')
    .eq('id', id)
    .maybeSingle();

  if (!ticket) return res.json({ status: 'INVALID' });

  if (ticket.status === 'USED') {
    return res.json({ status: 'ALREADY_USED' });
  }

  res.json({ status: 'VALID' });
});

// ==========================
// CONFIRM ENTRY
// ==========================
app.post('/check/:id/verify', async (req, res) => {
  const { id } = req.params;

  await supabase
    .from('tickets')
    .update({ status: 'USED' })
    .eq('id', id);

  res.json({ status: 'CONFIRMED' });
});

// ==========================
app.listen(3000, () => {
  console.log('Server running on port 3000');
});
