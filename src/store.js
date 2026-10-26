import { DB } from './db.js';
import { today, ACCOUNTS, TXN_CATEGORIES, HISTORICAL_EXPENSES, HISTORICAL_COGS, DEFAULT_CATEGORY_BIZ } from './constants.js';
import SEED_ITEMS    from './data/items.js';
import SEED_PARTIES  from './data/parties.js';
import LAST_RATES    from './data/lastRates.js';
import MODE_RATES    from './data/modeRates.js';
import PARTY_INVOICES_RAW from './data/partyInvoices.js';
import PARTY_HISTORY from './data/partyHistory.js';
import SUPPLIER_HISTORY from './data/supplierHistory.js';
import EXPENSE_SUBLEDGER from './data/expenseSubLedger.js';
import ACCOUNT_HISTORY from './data/accountHistory.js';

export const PARTY_INVOICES = PARTY_INVOICES_RAW;  // already a Map
export { LAST_RATES, MODE_RATES, PARTY_HISTORY, SUPPLIER_HISTORY, EXPENSE_SUBLEDGER, ACCOUNT_HISTORY };

export const _lastRateMap = new Map(LAST_RATES.map(([p,i,r]) => [p+'|'+i, r]));
export const _modeRateMap = new Map(MODE_RATES.map(([i,r]) => [i, r]));

export const ACCOUNT_OB = Object.fromEntries(ACCOUNTS.map(a => [a.id, a.ob]));

// Prefers an actual rate from a NEW in-app sale over the precomputed
// historical tables, since any in-app sale postdates the historical cutover
// and is a truer "what we actually charged last" than a frozen snapshot.
// itemId is optional — omitting it (older call sites) falls back to the
// original historical-only behavior.
export function getBestRate(pn, ii, ic, def, itemId) {
  let partyBest = null, anyBest = null;
  if (itemId) {
    for (const s of state.sales) {
      if (s.type === 'return') continue;
      for (const l of (s.items || [])) {
        if (l.itemId !== itemId || !l.rate) continue;
        if (s.partyName === pn && (!partyBest || s.date > partyBest.date)) partyBest = { date: s.date, rate: l.rate };
        if (!anyBest || s.date > anyBest.date) anyBest = { date: s.date, rate: l.rate };
      }
    }
  }
  if (partyBest) return partyBest.rate;
  if (_lastRateMap.has(pn+'|'+ii)) return _lastRateMap.get(pn+'|'+ii);
  if (_lastRateMap.has(pn+'|'+ic)) return _lastRateMap.get(pn+'|'+ic);
  if (_modeRateMap.has(ii)) return _modeRateMap.get(ii);
  if (anyBest) return anyBest.rate;
  return def || 0;
}

export function gPS(nm) {
  if (!nm) return null;
  const h = PARTY_HISTORY[nm];
  if (h?.s?.length) {
    const total = h.s.reduce((s,e) => s + (e[1]||0), 0);
    const months = Math.max(1, Math.round((new Date() - new Date(h.s[h.s.length-1][0]+'T00:00:00')) / (30*86400000)));
    return { ts: total, am: Math.round(total/months) };
  }
  return null;
}

export function getPetiFromName(n) {
  const nl = n.toLowerCase();
  if (nl.includes('mini body')) return 10;
  if (nl.includes('infrared')||nl.includes('full touch')||nl.includes('jaali')) return 6;
  return 1;
}

export function parseQty(raw, peti) {
  if (!raw) return NaN;
  const s = String(raw).trim().toLowerCase();
  const pm = s.match(/^(\d+(?:\.\d+)?)\s*(?:peti|p)$/); if (pm) return Number(pm[1]) * (peti||1);
  const nm = s.match(/^(\d+(\.\d+)?)$/); if (nm) return Number(nm[1]);
  return NaN;
}

export function parseResponse(text) {
  const t = text.toLowerCase();
  const tags = []; let followUpDate = null; let label = '';
  const today0 = new Date(); today0.setHours(0,0,0,0);
  const dayMap = {sunday:0,monday:1,tuesday:2,wednesday:3,thursday:4,friday:5,saturday:6};
  if (t.match(/paid|kar diya|transfer|done/))   { tags.push('paid'); label='Paid ✓'; }
  else if (t.match(/today|aaj/))                 { tags.push('promise'); followUpDate=today(); label='Pays today'; }
  else if (t.match(/tomorrow|kal/))              { const d=new Date(today0); d.setDate(d.getDate()+1); followUpDate=d.toISOString().slice(0,10); tags.push('promise'); label='Pays tomorrow'; }
  else {
    const dm = t.match(/sunday|monday|tuesday|wednesday|thursday|friday|saturday/);
    if (dm) { const target=dayMap[dm[0]],curr=today0.getDay(),diff=((target-curr)+7)%7||7,d=new Date(today0); d.setDate(d.getDate()+diff); followUpDate=d.toISOString().slice(0,10); tags.push('promise'); label='Pays '+dm[0]; }
    const nd = t.match(/(\d+)\s*(?:days?|din)/);
    if (nd) { const n=parseInt(nd[1]),d=new Date(today0); d.setDate(d.getDate()+n); followUpDate=d.toISOString().slice(0,10); tags.push('promise'); label='Pays in '+n+'d'; }
  }
  if (t.match(/busy|bahar/))      { tags.push('unavailable'); label=label||'Busy'; }
  if (t.match(/no pick|nahi/))    { tags.push('no-answer');   label=label||'No answer'; }
  if (t.match(/dispute|check/))   { tags.push('dispute');     label=label||'Disputed'; }
  if (!tags.length)                { tags.push('note');        label='Note'; }
  return { tags, followUpDate, label };
}

export function getAccBal(accId) {
  let bal = ACCOUNT_OB[accId] || 0;
  for (const t of state.txns) {
    if (t.type === 'transfer') {
      if (t.account === accId)   bal -= t.amount;
      if (t.toAccount === accId) bal += t.amount;
    } else if (t.account === accId) {
      bal += (t.type==='receipt'||t.type==='party_payment') ? t.amount : -t.amount;
    }
  }
  return bal;
}

export function getSupplierBal(partyId) {
  const party = state.parties.find(p => p.id === partyId);
  if (!party) return 0;
  // p.extraChargesTotal (loading/packing/octroi etc. named on the purchase
  // itself, billed by the same supplier) is owed to this same party, unlike
  // freight which is owed to a separate transporter — so it rolls in here.
  const purch   = state.purchases.filter(p => p.partyId === partyId).reduce((s,p) => s + p.total + (p.extraChargesTotal||0), 0);
  const freight = state.freightCharges.filter(f => f.transporterId === partyId).reduce((s,f) => s + f.amount, 0);
  const paid    = state.supplierPayments.filter(p => p.partyId === partyId).reduce((s,p) => s + p.amount, 0);
  return (party.ob || 0) + purch + freight - paid;
}

export function getCustomerBal(partyId) {
  const party = state.parties.find(p => p.id === partyId);
  if (!party) return 0;
  const sales = state.sales.filter(s => s.partyId === partyId).reduce((s,t) => s + t.total, 0);
  const paid  = state.payments.filter(p => p.partyId === partyId).reduce((s,p) => s + p.amount, 0);
  return (party.ob || 0) + sales - paid;
}

// FIFO bill queue for a customer: historical PARTY_INVOICES rows + new ERP sales,
// oldest-first, with all-time ERP payments (and return credits) consumed against
// the oldest outstanding bill first. `party.ob` for a party is the rolled-up sum of
// its historical invoices' outstanding amounts (verified 1:1 against seed data), so
// any excess is surfaced as a synthetic "Opening Balance" bill dated before all others
// — this keeps the queue correct even after a manual opening-balance edit.
export function getPartyBills(party) {
  if (!party) return [];
  // A merged-away duplicate party's historical invoices are filed under its
  // old name in the static import data — mergedNames (set by mergeParties)
  // is how the survivor keeps pulling that history in after the merge.
  const names = [party.name, ...(party.mergedNames||[])];
  const hist = names.flatMap(n => (PARTY_INVOICES.get(n) || []).map(([date,total,outstanding,,items]) => ({
    date, total, orig: outstanding, items, isNew: false,
  })));
  const histSum = hist.reduce((s,b) => s + b.orig, 0);
  const obExtra = (party.ob || 0) - histSum;
  if (obExtra > 0) hist.unshift({ date: '1900-01-01', total: obExtra, orig: obExtra, items: 'Opening Balance', isNew: false });

  const newBills = state.sales.filter(s => s.partyId === party.id && s.type !== 'return').map(s => ({
    date: s.date, total: s.total, orig: s.total,
    items: (s.items||[]).slice(0,3).map(x => (x.customerName||x.itemName)+' ×'+x.qty).join(', ') || 'Sale',
    isNew: true,
  }));
  const returnsTotal = Math.abs(state.sales.filter(s => s.partyId === party.id && s.type === 'return').reduce((s,t) => s + t.total, 0));
  const paymentsTotal = state.payments.filter(p => p.partyId === party.id).reduce((s,p) => s + p.amount, 0);

  const bills = [...hist, ...newBills].sort((a,b) => a.date.localeCompare(b.date));
  let rem = paymentsTotal + returnsTotal;
  const today0 = new Date();
  return bills.map((b,i) => {
    const applied = Math.min(rem, b.orig);
    rem = Math.max(0, rem - applied);
    const remaining = Math.max(0, b.orig - applied);
    const daysSince = Math.floor((today0 - new Date(b.date+'T00:00:00')) / 86400000);
    const overdue = Math.max(0, daysSince - (party.credit || 30));
    return { ...b, key: (b.isNew?'N':'H')+'_'+b.date+'_'+i, remaining, paid: b.orig>0 && remaining===0, overdue, daysSince };
  });
}

// FIFO bill queue for a supplier or transporter — same shape/contract as
// getPartyBills() above, so the Accounts payment form's invoice picker can
// render both with one code path. Unlike customers, suppliers/transporters
// have no per-invoice pre-app history (SUPPLIER_HISTORY only has aggregate
// totals), so the queue is just this app's own purchase/freight records —
// fine, since the point is to make GOING-FORWARD vendor/transporter payments
// bill-aware, not to reconstruct pre-app bills. party.ob (pre-app payable) is
// surfaced as its own synthetic oldest bill, same pattern as obExtra above.
export function getSupplierBills(party) {
  if (!party) return [];
  const isTransporterParty = party.type === 'transporter';
  const rawBills = isTransporterParty
    ? state.freightCharges.filter(f => f.transporterId === party.id).map(f => ({
        date: f.date, total: f.amount, orig: f.amount,
        items: f.notes || 'Freight', isNew: true,
      }))
    : state.purchases.filter(p => p.partyId === party.id && p.type !== 'return').map(p => ({
        date: p.date, total: p.total + (p.extraChargesTotal||0), orig: p.total + (p.extraChargesTotal||0),
        items: [(p.items||[]).slice(0,3).map(x => (x.supplierName||x.itemName)+' ×'+x.qty).join(', ') || 'Purchase',
          ...((p.extraCharges||[]).map(c => c.name+' ₹'+c.amount))].join(' · '),
        isNew: true,
      }));
  // Purchase returns (freight has no equivalent) are a credit against the
  // running payable, consumed FIFO against the oldest bill first — same
  // treatment sales returns get for customers in getPartyBills().
  const returnsTotal = isTransporterParty ? 0 : Math.abs(state.purchases.filter(p => p.partyId === party.id && p.type === 'return').reduce((s,t) => s + t.total, 0));
  const paymentsTotal = state.supplierPayments.filter(p => p.partyId === party.id).reduce((s,p) => s + p.amount, 0);

  const bills = rawBills.sort((a,b) => a.date.localeCompare(b.date));
  if (party.ob) bills.unshift({ date: '1900-01-01', total: party.ob, orig: party.ob, items: 'Opening Balance', isNew: false });

  let rem = paymentsTotal + returnsTotal;
  const today0 = new Date();
  return bills.map((b,i) => {
    const applied = Math.min(rem, Math.max(0, b.orig));
    rem = Math.max(0, rem - applied);
    const remaining = Math.max(0, b.orig - applied);
    const daysSince = Math.floor((today0 - new Date(b.date+'T00:00:00')) / 86400000);
    const overdue = Math.max(0, daysSince - (party.credit || 30));
    return { ...b, key: (b.isNew?'N':'H')+'_'+b.date+'_'+i, remaining, paid: b.orig>0 && remaining===0, overdue, daysSince };
  });
}

// Per-customer analytics: total sales, gross profit (using each item's CURRENT
// purchase rate as a cost proxy — the source Excel has no historical per-sale cost,
// so this is an approximation, not an exact historical margin), average days to
// collect payment (FIFO-matched oldest-bill-vs-oldest-payment, weighted by amount),
// and a grade based on that collection speed.
export function getAllCustomerMetrics() {
  const itemCost = new Map();
  for (const it of state.items) {
    if (it.internal) itemCost.set(it.internal.toLowerCase(), it.purch || 0);
    if (it.customer) itemCost.set((it.customer || '').toLowerCase(), it.purch || 0);
  }
  const costOf = name => itemCost.get(String(name || '').toLowerCase()) || 0;

  const customers = state.parties.filter(p => p.type === 'customer');
  return customers.map(p => {
    const hist = PARTY_HISTORY[p.name] || {};
    const histBills = (hist.s || []).map(([d, total, items]) => ({ date: d, total, items }));
    const newBills = state.sales.filter(s => s.partyId === p.id).map(s => ({
      date: s.date, total: s.total, items: (s.items || []).map(x => [x.customerName || x.itemName, x.qty, x.rate]),
    }));
    const bills = [...histBills, ...newBills];

    let totalSales = 0, totalCost = 0;
    for (const b of bills) {
      totalSales += b.total;
      for (const [name, qty] of b.items) totalCost += (qty || 0) * costOf(name);
    }
    const profit = Math.round(totalSales - totalCost);
    totalSales = Math.round(totalSales);
    const avgMarginPct = totalSales > 0 ? (profit / totalSales * 100) : 0;

    const payEvents = [
      ...(hist.p || []).map(([d, amt]) => ({ date: d, amt })),
      ...state.payments.filter(pm => pm.partyId === p.id).map(pm => ({ date: pm.date, amt: pm.amount })),
    ].sort((a, b) => a.date.localeCompare(b.date));
    const billQueue = [...bills].sort((a, b) => a.date.localeCompare(b.date)).map(b => ({ date: b.date, rem: b.total }));
    let bi = 0, weightedDaysSum = 0, matchedAmt = 0;
    for (const pay of payEvents) {
      let remaining = pay.amt;
      while (remaining > 0.01 && bi < billQueue.length) {
        const bill = billQueue[bi];
        if (bill.rem <= 0.01) { bi++; continue; }
        const applied = Math.min(remaining, bill.rem);
        const days = Math.max(0, (new Date(pay.date) - new Date(bill.date)) / 86400000);
        weightedDaysSum += days * applied; matchedAmt += applied;
        bill.rem -= applied; remaining -= applied;
        if (bill.rem <= 0.01) bi++;
      }
    }
    const avgPaymentDays = matchedAmt > 0 ? Math.round(weightedDaysSum / matchedAmt) : null;
    const currentOD = getCustomerBal(p.id);
    const grade = avgPaymentDays === null ? (currentOD > 0 ? 'D' : 'A')
      : avgPaymentDays <= 15 ? 'A' : avgPaymentDays <= 30 ? 'B' : avgPaymentDays <= 60 ? 'C' : 'D';

    return { party: p, totalSales, profit, avgMarginPct, avgPaymentDays, currentOD, grade, orderCount: bills.length };
  });
}

// Per-vendor analytics. Historical dated purchases AND payments come from
// SUPPLIER_HISTORY (raw Purchase sheet's Cash Paid + Bank Paid columns), so avg
// payment days is FIFO-matched across the full history, not just new ERP data.
export function getAllVendorMetrics() {
  const suppliers = state.parties.filter(p => p.type === 'supplier');
  return suppliers.map(p => {
    const hist = SUPPLIER_HISTORY[p.name] || {};
    const histBills = (hist.s || []).map(([d, total]) => ({ date: d, total }));
    const newBills = state.purchases.filter(pu => pu.partyId === p.id).map(pu => ({ date: pu.date, total: pu.total + (pu.extraChargesTotal||0) }));
    const totalPurchased = Math.round([...histBills, ...newBills].reduce((s, b) => s + b.total, 0));

    const payEvents = [
      ...(hist.p || []).map(([d, amt]) => ({ date: d, amt })),
      ...state.supplierPayments.filter(sp => sp.partyId === p.id).map(sp => ({ date: sp.date, amt: sp.amount })),
    ].sort((a, b) => a.date.localeCompare(b.date));
    const billQueue = [...histBills, ...newBills].sort((a, b) => a.date.localeCompare(b.date)).map(b => ({ date: b.date, rem: b.total }));
    let bi = 0, weightedDaysSum = 0, matchedAmt = 0;
    for (const pay of payEvents) {
      let remaining = pay.amt;
      while (remaining > 0.01 && bi < billQueue.length) {
        const bill = billQueue[bi];
        if (bill.rem <= 0.01) { bi++; continue; }
        const applied = Math.min(remaining, bill.rem);
        const days = Math.max(0, (new Date(pay.date) - new Date(bill.date)) / 86400000);
        weightedDaysSum += days * applied; matchedAmt += applied;
        bill.rem -= applied; remaining -= applied;
        if (bill.rem <= 0.01) bi++;
      }
    }
    const avgPaymentDays = matchedAmt > 0 ? Math.round(weightedDaysSum / matchedAmt) : null;
    const currentPayable = getSupplierBal(p.id);
    const grade = avgPaymentDays === null ? null
      : avgPaymentDays <= 7 ? 'A' : avgPaymentDays <= 15 ? 'B' : avgPaymentDays <= 30 ? 'C' : 'D';

    return { party: p, totalPurchased, orderCount: histBills.length + newBills.length, avgPaymentDays, currentPayable, grade };
  });
}

// ── Profit Analysis: per-sale FIFO profit + interest-load (1%/month) ───────
// Reuses the exact FIFO bill-vs-payment-event matching getAllCustomerMetrics/
// getAllVendorMetrics already do (weightedDaysSum/matchedAmt there) — but
// per-BILL (not aggregated per party) and folding any still-unpaid remainder
// into "settled as of today" rather than leaving it unmatched. That fold is
// a deliberate modelling choice, not an oversight: if a bill is still open,
// the capital (or the credit) is still live right now, so the interest
// clock for it is still running — using today() here is what makes the
// resulting number a live, moving figure rather than a frozen one that
// ignores open bills entirely. Returns a Map keyed by the bill's own id
// (purchase id for the supplier side, sale id for the customer side) to a
// weighted-average settlement date, in epoch ms.
const msOf = dateStr => new Date((dateStr || today()) + 'T00:00:00').getTime();

export function getPurchasePaidDateMap() {
  const out = new Map();
  const todayMs = msOf(today());
  for (const p of state.parties.filter(p => p.type === 'supplier')) {
    const hist = SUPPLIER_HISTORY[p.name] || {};
    const histBills = (hist.s || []).map(([d, total]) => ({ id: null, date: d, total }));
    const newBills = state.purchases.filter(pu => pu.partyId === p.id && pu.type !== 'return')
      .map(pu => ({ id: pu.id, date: pu.date, total: pu.total + (pu.extraChargesTotal || 0) }));
    const payEvents = [
      ...(hist.p || []).map(([d, amt]) => ({ date: d, amt })),
      ...state.supplierPayments.filter(sp => sp.partyId === p.id).map(sp => ({ date: sp.date, amt: sp.amount })),
    ].sort((a, b) => a.date.localeCompare(b.date));
    const billQueue = [...histBills, ...newBills].sort((a, b) => a.date.localeCompare(b.date))
      .map(b => ({ ...b, rem: b.total, weightedSum: 0, matched: 0 }));
    let bi = 0;
    for (const pay of payEvents) {
      let remaining = pay.amt;
      while (remaining > 0.01 && bi < billQueue.length) {
        const bill = billQueue[bi];
        if (bill.rem <= 0.01) { bi++; continue; }
        const applied = Math.min(remaining, bill.rem);
        bill.weightedSum += msOf(pay.date) * applied; bill.matched += applied;
        bill.rem -= applied; remaining -= applied;
        if (bill.rem <= 0.01) bi++;
      }
    }
    for (const bill of billQueue) {
      if (!bill.id) continue; // only ERP purchases are traceable back to a sale's FIFO cost layer
      if (bill.rem > 0.01) { bill.weightedSum += todayMs * bill.rem; bill.matched += bill.rem; }
      out.set(bill.id, bill.matched > 0 ? bill.weightedSum / bill.matched : todayMs);
    }
  }
  return out;
}

export function getSaleReceivedDateMap() {
  const out = new Map();
  const todayMs = msOf(today());
  for (const p of state.parties.filter(p => p.type === 'customer')) {
    const hist = PARTY_HISTORY[p.name] || {};
    const histBills = (hist.s || []).map(([d, total]) => ({ id: null, date: d, total }));
    const newBills = state.sales.filter(s => s.partyId === p.id && s.type !== 'return').map(s => ({ id: s.id, date: s.date, total: s.total }));
    const payEvents = [
      ...(hist.p || []).map(([d, amt]) => ({ date: d, amt })),
      ...state.payments.filter(pm => pm.partyId === p.id).map(pm => ({ date: pm.date, amt: pm.amount })),
      // Sales returns are a same-day credit against the oldest open bill —
      // same treatment getPartyBills gives them — fed into the same FIFO
      // match as if they were a payment dated on the return itself.
      ...state.sales.filter(s => s.partyId === p.id && s.type === 'return').map(s => ({ date: s.date, amt: Math.abs(s.total) })),
    ].sort((a, b) => a.date.localeCompare(b.date));
    const billQueue = [...histBills, ...newBills].sort((a, b) => a.date.localeCompare(b.date))
      .map(b => ({ ...b, rem: b.total, weightedSum: 0, matched: 0 }));
    let bi = 0;
    for (const pay of payEvents) {
      let remaining = pay.amt;
      while (remaining > 0.01 && bi < billQueue.length) {
        const bill = billQueue[bi];
        if (bill.rem <= 0.01) { bi++; continue; }
        const applied = Math.min(remaining, bill.rem);
        bill.weightedSum += msOf(pay.date) * applied; bill.matched += applied;
        bill.rem -= applied; remaining -= applied;
        if (bill.rem <= 0.01) bi++;
      }
    }
    for (const bill of billQueue) {
      if (!bill.id) continue; // only ERP sales get a per-sale profit row
      if (bill.rem > 0.01) { bill.weightedSum += todayMs * bill.rem; bill.matched += bill.rem; }
      out.set(bill.id, bill.matched > 0 ? bill.weightedSum / bill.matched : todayMs);
    }
  }
  return out;
}

// Per-sale FIFO profit — same layer replay as getFifoSaleMarginMap, but (a)
// kept per-sale rather than aggregated per-item, (b) tags every layer with
// the purchase id/date that supplied it, and (c) looks up that purchase's
// paid-date and the sale's own received-date (both from the two functions
// above) to compute a 1%/month interest load: the financing cost of the gap
// between when cash left for the stock and when cash came back in for
// selling it. A negative interest load is a real, meaningful result — it
// means the customer's payment arrived before the supplier had to be paid,
// i.e. this particular deal funded itself and actually earned a float
// benefit rather than costing one.
//
// Cost drawn from stock whose origin isn't a logged in-app purchase (the
// opening-stock/baseline reconciliation layer, or a sales-return put back
// at current cost) has no purchase to date-match against, so it's excluded
// from the interest-bearing base (`traceableCost`) — `cost`/`grossProfit`
// still include it, only the interest calc does not. This is called out in
// the UI rather than silently assumed.
export function getSaleProfitAnalysis() {
  const eventsById = {};
  const add = (id, date, seq, type, netQty, rate, meta) => {
    if (!id) return;
    (eventsById[id] ||= []).push({ date, seq: seq || '', type, netQty, rate, meta });
  };
  for (const p of state.purchases) for (const l of (p.items || [])) add(l.itemId, p.date, p.createdAt, 'purch', l.qty || 0, l.rate || 0, { purchaseId: p.id });
  for (const s of state.sales) for (const x of (s.items || [])) add(x.itemId, s.date, s.createdAt, 'sale', -(x.qty || 0), x.rate, { saleId: s.id });
  for (const r of state.repairExpenses) add(r.itemId, r.date, r.createdAt, 'repair', -(r.qty || 0), null, {});

  const bySale = new Map(); // saleId -> { qty, revenue, cost, segs:[{purchaseId,amt}] }
  for (const it of state.items) {
    const events = (eventsById[it.id] || []).slice().sort((a, b) => a.date.localeCompare(b.date) || a.seq.localeCompare(b.seq));
    const netFromEvents = events.reduce((s, e) => s + e.netQty, 0);
    const plug = (it.stock || 0) - netFromEvents;
    const layers = [];
    if (plug > 1e-9) layers.push({ qty: plug, rate: it.purch || 0, purchaseId: null });
    const withdrawDetailed = qty => {
      let need = qty; const segs = [];
      while (need > 1e-9 && layers.length) {
        const l = layers[0]; const take = Math.min(l.qty, need);
        segs.push({ qty: take, rate: l.rate, purchaseId: l.purchaseId });
        l.qty -= take; need -= take;
        if (l.qty <= 1e-9) layers.shift();
      }
      if (need > 1e-9) segs.push({ qty: need, rate: it.purch || 0, purchaseId: null });
      return segs;
    };
    for (const e of events) {
      if (e.type === 'purch') {
        if (e.netQty > 1e-9) layers.push({ qty: e.netQty, rate: e.rate, purchaseId: e.meta.purchaseId });
        else if (e.netQty < -1e-9) withdrawDetailed(-e.netQty);
      } else if (e.type === 'repair') {
        if (e.netQty < -1e-9) withdrawDetailed(-e.netQty);
      } else if (e.type === 'sale') {
        if (e.netQty < -1e-9) {
          const qty = -e.netQty;
          const segs = withdrawDetailed(qty);
          const cost = segs.reduce((s, g) => s + g.qty * g.rate, 0);
          const rec = bySale.get(e.meta.saleId) || { qty: 0, revenue: 0, cost: 0, segs: [] };
          rec.qty += qty; rec.revenue += qty * (e.rate || 0); rec.cost += cost;
          rec.segs.push(...segs.map(g => ({ purchaseId: g.purchaseId, amt: g.qty * g.rate })));
          bySale.set(e.meta.saleId, rec);
        } else if (e.netQty > 1e-9) {
          layers.push({ qty: e.netQty, rate: it.purch || 0, purchaseId: null }); // sales return, back at current cost
        }
      }
    }
  }

  const purchasePaidMap = getPurchasePaidDateMap();
  const saleReceivedMap = getSaleReceivedDateMap();
  const todayMs = msOf(today());
  const out = [];
  for (const s of state.sales) {
    if (s.type === 'return') continue;
    const rec = bySale.get(s.id);
    if (!rec) continue;
    const party = state.parties.find(p => p.id === s.partyId);
    const grossProfit = rec.revenue - rec.cost;
    const grossMarginPct = rec.revenue > 0 ? Math.round(grossProfit / rec.revenue * 1000) / 10 : 0;
    const traceableSegs = rec.segs.filter(g => g.purchaseId);
    const traceableCost = traceableSegs.reduce((s2, g) => s2 + g.amt, 0);
    let interestLoad = 0;
    if (traceableCost > 1e-6) {
      const weightedPurchMs = traceableSegs.reduce((s2, g) => s2 + (purchasePaidMap.get(g.purchaseId) ?? todayMs) * g.amt, 0) / traceableCost;
      const saleRecMs = saleReceivedMap.get(s.id) ?? todayMs;
      const months = (saleRecMs - weightedPurchMs) / (1000 * 60 * 60 * 24 * 30);
      interestLoad = Math.round(traceableCost * 0.01 * months);
    }
    const netProfit = Math.round(grossProfit - interestLoad);
    const netMarginPct = rec.revenue > 0 ? Math.round(netProfit / rec.revenue * 1000) / 10 : 0;
    out.push({
      saleId: s.id, date: s.date, partyId: s.partyId, partyName: party?.name || s.partyName || '—',
      qty: Math.round(rec.qty * 100) / 100, revenue: Math.round(rec.revenue), cost: Math.round(rec.cost),
      grossProfit: Math.round(grossProfit), grossMarginPct,
      traceableCost: Math.round(traceableCost), interestLoad, netProfit, netMarginPct,
    });
  }
  return out.sort((a, b) => b.date.localeCompare(a.date));
}

// Per-customer roll-up of getSaleProfitAnalysis — avg profit % (both gross
// and interest-adjusted) and total interest load per customer, ranked by
// net margin so the "where am I earning more / losing because of interest
// load" comparison (request #4) is a straight sort, not a separate calc.
export function getCustomerProfitSummary() {
  const sales = getSaleProfitAnalysis();
  const byParty = new Map();
  for (const s of sales) {
    const cur = byParty.get(s.partyId) || { partyId: s.partyId, partyName: s.partyName, revenue: 0, grossProfit: 0, interestLoad: 0, netProfit: 0, orderCount: 0 };
    cur.revenue += s.revenue; cur.grossProfit += s.grossProfit; cur.interestLoad += s.interestLoad; cur.netProfit += s.netProfit; cur.orderCount++;
    byParty.set(s.partyId, cur);
  }
  return [...byParty.values()].map(c => ({
    ...c,
    grossMarginPct: c.revenue > 0 ? Math.round(c.grossProfit / c.revenue * 1000) / 10 : 0,
    netMarginPct: c.revenue > 0 ? Math.round(c.netProfit / c.revenue * 1000) / 10 : 0,
    interestLoadPctOfRevenue: c.revenue > 0 ? Math.round(c.interestLoad / c.revenue * 1000) / 10 : 0,
  })).sort((a, b) => b.netMarginPct - a.netMarginPct);
}

export function getCategoryBiz(catId) {
  if (state.categoryTags[catId]) return state.categoryTags[catId];
  return DEFAULT_CATEGORY_BIZ[catId] || 'business';
}

export function setCategoryBiz(catId, tag) {
  state.categoryTags = { ...state.categoryTags, [catId]: tag };
  save.categoryTags();
}

// Expense categories = every "out" category with no party linkage (plain expense
// categories + any custom ones) — excludes Vendor/Transporter Payment (those are
// payables, handled via COGS/payables, not operating expenses) and Transfer (contra).
export function getExpenseReport() {
  const cats = [...TXN_CATEGORIES, ...state.customCategories].filter(c => c.dir === 'out' && c.party === null);
  return cats.map(c => {
    const historical = HISTORICAL_EXPENSES[c.id] || 0;
    const newAmt = state.txns.filter(t => t.type === 'payment' && t.category === c.id).reduce((s, t) => s + t.amount, 0);
    return { id: c.id, label: c.label, historical, newAmt, total: historical + newAmt, biz: getCategoryBiz(c.id) };
  }).sort((a, b) => b.total - a.total);
}

// Sub-ledger within a category: groups entries by the individual/head they belong
// to (e.g. each staff member under Salary) — combining dated historical entries
// mined from the Bank Entries sheet with new ERP entries sharing the same payee.
// Categories with no historical breakdown (not every category has one) just show
// whatever's been entered in this app.
export function getCategorySubLedger(catId) {
  const histSubs = EXPENSE_SUBLEDGER[catId] || {};
  const groups = {};
  for (const name in histSubs) groups[name] = { name, histEntries: histSubs[name], newEntries: [] };

  const newTxns = state.txns.filter(t => t.type === 'payment' && t.category === catId);
  for (const t of newTxns) {
    const raw = (t.payee || t.partyName || 'Other').trim() || 'Other';
    const key = Object.keys(groups).find(k => k.toLowerCase() === raw.toLowerCase()) || raw;
    if (!groups[key]) groups[key] = { name: key, histEntries: [], newEntries: [] };
    groups[key].newEntries.push(t);
  }

  return Object.values(groups).map(g => {
    const histTotal = g.histEntries.reduce((s, e) => s + e[1], 0);
    const newTotal = g.newEntries.reduce((s, t) => s + t.amount, 0);
    return { ...g, histTotal, newTotal, total: histTotal + newTotal };
  }).sort((a, b) => b.total - a.total);
}

// Some historical rows record a sale and a return together under one dated
// entry with a single blended total (e.g. goods sold plus an unrelated return
// processed the same day). Split those into a clean sale segment and a clean
// return segment instead of showing one row whose sign and item list are a
// confusing mix — but leave a normal single-direction row untouched so its
// recorded total (which may include adjustments not reconstructable from the
// item lines) isn't second-guessed.
export function splitPosNegItems(total,items){
  const arr=items||[];
  const pos=arr.filter(x=>Array.isArray(x)?x[1]>0:true);
  const neg=arr.filter(x=>Array.isArray(x)&&x[1]<0);
  if(pos.length&&neg.length){
    const posTotal=pos.reduce((s,x)=>s+x[1]*x[2],0);
    const negTotal=neg.reduce((s,x)=>s+x[1]*x[2],0);
    return[{total:posTotal,items:pos,isReturn:false},{total:negTotal,items:neg,isReturn:true}];
  }
  return[{total,items:arr,isReturn:total<0}];
}

// Flat log of every item movement (historical + ERP), sale and purchase side.
// Historical rows carry no itemId (Excel data only has item names), so matching
// to a specific item falls back to case-insensitive name comparison; ERP rows
// carry itemId directly since SaleEntry/PurchaseEntry stamp it at save time.
export function getAllItemRows() {
  const rows = [];
  for (const [party, hist] of Object.entries(PARTY_HISTORY)) {
    for (const s of (hist.s || [])) {
      const [d, total, items] = s;
      // A historical return invoice may record its items with a positive
      // per-item qty (only the invoice TOTAL carries the negative sign) — normalize
      // to a negative qty/amount here so downstream consumers (Item Log, per-item
      // sold/revenue totals) can treat sign and the isReturn flag consistently,
      // the same way a return recorded in-app already does.
      for (const seg of splitPosNegItems(total, items)) {
        for (const it of (seg.items || [])) {
          const [name, rawQty, rate] = Array.isArray(it) ? it : [it, 1, 0];
          const qty = seg.isReturn ? -Math.abs(rawQty||0) : Math.abs(rawQty||0);
          rows.push({ d, item: name, qty, rate, amount: qty*(rate||0), dir: 'out', party, isReturn: seg.isReturn });
        }
      }
    }
  }
  for (const s of state.sales) {
    const isReturn = s.type === 'return';
    for (const x of (s.items || [])) {
      // saleId lets a per-row consumer (e.g. the Item Ledger's margin column)
      // look this exact transaction up in getFifoSaleCosts(itemId) below —
      // historical rows elsewhere in this function have no equivalent id,
      // same documented exclusion getFifoStockMap/getFifoSaleMarginMap use.
      rows.push({ d: s.date, item: x.customerName||x.itemName, itemId: x.itemId, qty: x.qty, rate: x.rate, amount: x.amount ?? ((x.qty||0)*(x.rate||0)), dir: 'out', party: s.partyName, isNew: true, isReturn, saleId: s.id });
    }
  }
  for (const [vendor, hist] of Object.entries(SUPPLIER_HISTORY)) {
    for (const s of (hist.s || [])) {
      const [d, , items] = s;
      for (const it of (items || [])) {
        const [name, qty, rate] = Array.isArray(it) ? it : [it, 1, 0];
        rows.push({ d, item: name, qty, rate, amount: (qty||0)*(rate||0), dir: 'in', party: vendor });
      }
    }
  }
  for (const p of state.purchases) {
    for (const x of (p.items || [])) {
      rows.push({ d: p.date, item: x.customerName||x.itemName, itemId: x.itemId, qty: x.qty, rate: x.rate, amount: (x.qty||0)*(x.rate||0), dir: 'in', party: p.partyName, isNew: true });
    }
  }
  return rows;
}

// Per-item metrics: pulls every sale/purchase row that matches this item (by
// itemId for ERP rows, by internal/customer name for historical Excel rows,
// since those predate itemId). mergedNames (set by mergeItems when a
// duplicate item is folded into this one) is included so historical rows
// filed under the absorbed item's old name keep counting toward it. COGS uses
// the item's *current* purchase rate as a proxy — same convention as getPnL —
// since historical per-sale cost isn't in the source data.
export function getItemMetrics(item, allRows) {
  const rows = allRows || getAllItemRows();
  const names = new Set([item.internal, item.name, item.customer, ...(item.mergedNames||[])].filter(Boolean).map(s => s.toLowerCase().trim()));
  const matched = rows.filter(r => (r.itemId && r.itemId === item.id) || (!r.itemId && names.has((r.item||'').toLowerCase().trim())));
  const soldRows = matched.filter(r => r.dir === 'out' && !r.isReturn);
  const returnRows = matched.filter(r => r.dir === 'out' && r.isReturn);
  const purchRows = matched.filter(r => r.dir === 'in');
  // Net returned qty/revenue back out of gross sold — a returned unit isn't
  // really "sold" any more, and counting it as a sale (as getAllItemRows used
  // to for historical returns, since it discarded the invoice's negative
  // total) inflated both totalSoldQty and totalRevenue.
  const totalSoldQty = Math.round(soldRows.reduce((s, r) => s + Math.abs(r.qty||0), 0) - returnRows.reduce((s, r) => s + Math.abs(r.qty||0), 0));
  const totalPurchQty = Math.round(purchRows.reduce((s, r) => s + Math.abs(r.qty||0), 0));
  const totalRevenue = soldRows.reduce((s, r) => s + Math.abs(r.amount||0), 0) - returnRows.reduce((s, r) => s + Math.abs(r.amount||0), 0);
  const totalCost = totalSoldQty * (item.purch||0);
  const profit = totalRevenue - totalCost;
  const avgProfitPct = totalRevenue > 0 ? Math.round(profit/totalRevenue*1000)/10 : 0;
  return { rows: matched, soldRows, purchRows, totalSoldQty, totalPurchQty, totalRevenue, totalCost, profit, avgProfitPct };
}

// All active items' metrics in one pass, sharing a single getAllItemRows() call
// so the per-item filtering (O(items × rows)) only happens once per render.
export function getAllItemMetrics() {
  const allRows = getAllItemRows();
  return state.items.filter(it => it.active !== false).map(it => ({ item: it, ...getItemMetrics(it, allRows) }));
}

// Profit & Loss: Revenue (all customer sales, historical + ERP) minus COGS (each
// item's current purchase rate × qty sold, a proxy since historical per-sale cost
// isn't in the source data) gives Gross Profit; minus Business-tagged operating
// expenses gives Net Profit. Non-business (personal/Ghar Kharch) spend is shown
// separately, not netted against business profit.
export function getPnL() {
  const custMetrics = getAllCustomerMetrics();
  const revenue = Math.round(custMetrics.reduce((s, r) => s + r.totalSales, 0));
  // COGS: the historical portion uses the real, transaction-by-transaction
  // Total COS figure (HISTORICAL_COGS) rather than today's purchase rate
  // applied across years of sales — that proxy overstated gross profit by
  // ~₹3.8M. Both HISTORICAL_COGS and the new-sales proxy below use the raw
  // purchase rate (item.purchRaw), not item.purch (which is landed cost
  // including an allocated transport share, kept that way deliberately for
  // margin display) — transport already has its own Expenses line, so using
  // landed cost here would double-count it.
  const itemCost = new Map();
  for (const it of state.items) {
    const rate = it.purchRaw ?? it.purch ?? 0;
    if (it.internal) itemCost.set(it.internal.toLowerCase(), rate);
    if (it.customer) itemCost.set((it.customer || '').toLowerCase(), rate);
  }
  const costOf = name => itemCost.get(String(name || '').toLowerCase()) || 0;
  let newCogs = 0;
  for (const s of state.sales) {
    for (const x of (s.items || [])) newCogs += (x.qty || 0) * costOf(x.customerName || x.itemName);
  }
  const cogs = Math.round(HISTORICAL_COGS + newCogs);
  const grossProfit = revenue - cogs;
  const expenseRows = getExpenseReport();
  const businessExpenses = expenseRows.filter(e => e.biz === 'business').reduce((s, e) => s + e.total, 0);
  const nonBusinessExpenses = expenseRows.filter(e => e.biz !== 'business').reduce((s, e) => s + e.total, 0);
  // Repair Expenditure — warranty cost incurred per party (PCB swaps, or
  // reimbursing a party for a repair they paid for themselves) — kept as its
  // own P&L line rather than folded into businessExpenses' category rows
  // above, since it's tracked per-party (repairExpenses) rather than by the
  // party-less TXN_CATEGORIES those rows are built from.
  const repairExpenseTotal = Math.round(state.repairExpenses.reduce((s, r) => s + (r.expenseAmount || 0), 0));
  const netProfit = grossProfit - businessExpenses - repairExpenseTotal;
  return { revenue, cogs, grossProfit, businessExpenses, nonBusinessExpenses, repairExpenseTotal, netProfit, expenseRows };
}

// Total warranty/repair expenditure booked against one party — sums both
// standalone PCB-return log entries and payment-time reimbursements.
export function getRepairExpenseTotal(partyId) {
  return state.repairExpenses.filter(r => r.partyId === partyId).reduce((s, r) => s + (r.expenseAmount || 0), 0);
}

// FIFO stock valuation — item.purch/item.sale are a single flat rate per item
// that's only ever refreshed by re-running the offline seed/import script, so
// as soon as a purchase is entered in-app at a different rate, item.purch
// silently goes stale and every stockValue computed from it (Balance Sheet,
// Category Analytics, Item Detail) is wrong. FIFO layers fix that by costing
// remaining stock from the actual sequence of purchases, but only using
// ITEM-ID-tagged, dated, in-app records (purchases/purchase-returns,
// sales/sales-returns, repair-expense consumption) as real cost layers —
// walking the bulk-imported historical rows (name-matched, not itemId-matched)
// was tried and rejected: replaying SUPPLIER_HISTORY/PARTY_HISTORY alone
// reconstructs a total quantity wildly different from item.stock for most
// items (verified: 3727 pcs vs 1410 pcs tracked, on real data), because that
// seed data is incomplete/inconsistent for this purpose — it was only ever
// meant to drive name-matched sales/profit metrics, not a per-unit cost
// ledger. So: replay only the dated in-app events for this item, oldest
// first; whatever portion of the item's CURRENT tracked stock isn't
// explained by those dated events (because it predates the app, or moved
// through an undated path like a damage/repair correction or the Move Stock
// tool) becomes one "baseline" layer valued at the item's stored purch rate,
// inserted at the very front (oldest, consumed first) — the same "opening
// balance" convention already used for party ledgers elsewhere in this app.
// A fresh purchase then immediately creates its own correctly-rated layer,
// and a sale/repair draws down the oldest layer first, so valuation updates
// automatically instead of requiring item.purch to be hand-edited.
export function getFifoStockMap() {
  const eventsById = {};
  const add = (id, date, seq, netQty, rate) => {
    if (!id) return;
    (eventsById[id] ||= []).push({ date, seq: seq || '', netQty, rate });
  };
  for (const p of state.purchases) {
    for (const l of (p.items || [])) {
      // l.qty is already signed: positive for a normal purchase line,
      // negative for a purchase-return line — l.rate is the real entered
      // cost, used whenever this line pushes a new layer (netQty>0).
      add(l.itemId, p.date, p.createdAt, l.qty || 0, l.rate || 0);
    }
  }
  for (const s of state.sales) {
    for (const x of (s.items || [])) {
      // x.qty is signed the same way: positive for a normal sale line
      // (stock leaves, netQty<0), negative for a sales-return line (stock
      // comes back, netQty>0). Rate is resolved per-item below at
      // it.purch — a returned unit is valued at current cost, not the price
      // it was billed at (billed price includes margin, not a valid cost).
      add(x.itemId, s.date, s.createdAt, -(x.qty || 0), null);
    }
  }
  for (const r of state.repairExpenses) {
    // only ever withdraws stock (netQty<0) — rate is never read for that branch.
    add(r.itemId, r.date, r.createdAt, -(r.qty || 0), null);
  }

  const map = {};
  for (const it of state.items) {
    const events = (eventsById[it.id] || []).slice().sort((a, b) => a.date.localeCompare(b.date) || a.seq.localeCompare(b.seq));
    const layers = [];
    const withdraw = qty => {
      let need = qty;
      while (need > 1e-9 && layers.length) {
        const l = layers[0];
        const take = Math.min(l.qty, need);
        l.qty -= take; need -= take;
        if (l.qty <= 1e-9) layers.shift();
      }
    };
    for (const e of events) {
      if (e.netQty > 1e-9) layers.push({ qty: e.netQty, rate: e.rate != null ? e.rate : (it.purch || 0) });
      else if (e.netQty < -1e-9) withdraw(-e.netQty);
    }
    // Reconcile to item.stock (the authoritative, incrementally-maintained
    // quantity — also touched by undated paths like damage moves and Move
    // Stock, which dated events above can't see). Any shortfall becomes the
    // oldest layer; any dated-side overshoot is trimmed oldest-first too.
    const inAppQty = layers.reduce((s, l) => s + l.qty, 0);
    const plug = (it.stock || 0) - inAppQty;
    if (plug > 1e-9) layers.unshift({ qty: plug, rate: it.purch || 0 });
    else if (plug < -1e-9) withdraw(-plug);
    const qty = Math.round(layers.reduce((s, l) => s + l.qty, 0) * 100) / 100;
    const value = Math.round(layers.reduce((s, l) => s + l.qty * l.rate, 0) * 100) / 100;
    const avgRate = qty > 0 ? value / qty : (it.purch || 0);
    map[it.id] = { qty, value, avgRate };
  }
  return map;
}

// Real, FIFO-costed margin on actual in-app sales — same dated/itemId-tagged
// event universe as getFifoStockMap (purchases, sales, repair-expense draws),
// replayed in chronological order per item so each sale line is costed from
// whatever layers were actually on hand *at that point in time*, not today's
// average. Bulk historical (pre-app, name-matched) sales are excluded — same
// restriction getFifoStockMap documents above: that seed data can't support a
// trustworthy per-unit cost reconstruction. Returns one map keyed by item id;
// soldQty/marginPct/marginPerPc are null-ish (0/null) for an item with no
// in-app sales yet, so callers can fall back to the theoretical
// sale-rate-vs-current-cost estimate.
export function getFifoSaleMarginMap() {
  const eventsById = {};
  const add = (id, date, seq, type, netQty, rate) => {
    if (!id) return;
    (eventsById[id] ||= []).push({ date, seq: seq || '', type, netQty, rate });
  };
  for (const p of state.purchases) for (const l of (p.items || [])) add(l.itemId, p.date, p.createdAt, 'purch', l.qty || 0, l.rate || 0);
  for (const s of state.sales) for (const x of (s.items || [])) add(x.itemId, s.date, s.createdAt, 'sale', -(x.qty || 0), x.rate);
  for (const r of state.repairExpenses) add(r.itemId, r.date, r.createdAt, 'repair', -(r.qty || 0), null);

  const map = {};
  for (const it of state.items) {
    const events = (eventsById[it.id] || []).slice().sort((a, b) => a.date.localeCompare(b.date) || a.seq.localeCompare(b.seq));
    const netFromEvents = events.reduce((s, e) => s + e.netQty, 0);
    const plug = (it.stock || 0) - netFromEvents; // same baseline-layer reconciliation as getFifoStockMap
    const layers = [];
    if (plug > 1e-9) layers.push({ qty: plug, rate: it.purch || 0 });
    const withdraw = qty => {
      let need = qty, costSum = 0, got = 0;
      while (need > 1e-9 && layers.length) {
        const l = layers[0]; const take = Math.min(l.qty, need);
        costSum += take * l.rate; got += take; l.qty -= take; need -= take;
        if (l.qty <= 1e-9) layers.shift();
      }
      if (need > 1e-9) { costSum += need * (it.purch || 0); got += need; } // ran out of dated layers — fall back to flat purch rate for the shortfall
      return got > 0 ? costSum / got : (it.purch || 0);
    };
    let soldQty = 0, soldRevenue = 0, soldCost = 0;
    for (const e of events) {
      if (e.type === 'purch') {
        if (e.netQty > 1e-9) layers.push({ qty: e.netQty, rate: e.rate });
        else if (e.netQty < -1e-9) withdraw(-e.netQty);
      } else if (e.type === 'repair') {
        if (e.netQty < -1e-9) withdraw(-e.netQty);
      } else if (e.type === 'sale') {
        if (e.netQty < -1e-9) {
          const qty = -e.netQty; const costRate = withdraw(qty);
          soldQty += qty; soldRevenue += qty * (e.rate || 0); soldCost += qty * costRate;
        } else if (e.netQty > 1e-9) {
          layers.push({ qty: e.netQty, rate: it.purch || 0 }); // sales return — back at current cost, not the billed price
        }
      }
    }
    const marginPerPc = soldQty > 0 ? Math.round((soldRevenue - soldCost) / soldQty * 100) / 100 : null;
    const marginPct = soldRevenue > 0 ? Math.round((soldRevenue - soldCost) / soldRevenue * 1000) / 10 : null;
    map[it.id] = { soldQty: Math.round(soldQty * 100) / 100, marginPerPc, marginPct };
  }
  return map;
}

// Same FIFO layer replay as getFifoSaleMarginMap, but filtered to a date
// range for revenue/cost accumulation while still replaying EVERY event
// (including purchases/sales outside the range) in chronological order —
// a sale's correct FIFO cost depends on the full layer history up to that
// point, not just what happened inside the window. Used by the Dashboard's
// "Profit This Month" tile (request #7) so it reflects true landed cost
// rather than the flat item.purch proxy.
export function getFifoProfitInRange(fromDate, toDate) {
  const eventsById = {};
  const add = (id, date, seq, type, netQty, rate) => {
    if (!id) return;
    (eventsById[id] ||= []).push({ date, seq: seq || '', type, netQty, rate });
  };
  for (const p of state.purchases) for (const l of (p.items || [])) add(l.itemId, p.date, p.createdAt, 'purch', l.qty || 0, l.rate || 0);
  for (const s of state.sales) for (const x of (s.items || [])) add(x.itemId, s.date, s.createdAt, 'sale', -(x.qty || 0), x.rate);
  for (const r of state.repairExpenses) add(r.itemId, r.date, r.createdAt, 'repair', -(r.qty || 0), null);

  let revenue = 0, cost = 0, qty = 0;
  for (const it of state.items) {
    const events = (eventsById[it.id] || []).slice().sort((a, b) => a.date.localeCompare(b.date) || a.seq.localeCompare(b.seq));
    const netFromEvents = events.reduce((s, e) => s + e.netQty, 0);
    const plug = (it.stock || 0) - netFromEvents;
    const layers = [];
    if (plug > 1e-9) layers.push({ qty: plug, rate: it.purch || 0 });
    const withdraw = q => {
      let need = q, costSum = 0, got = 0;
      while (need > 1e-9 && layers.length) {
        const l = layers[0]; const take = Math.min(l.qty, need);
        costSum += take * l.rate; got += take; l.qty -= take; need -= take;
        if (l.qty <= 1e-9) layers.shift();
      }
      if (need > 1e-9) { costSum += need * (it.purch || 0); got += need; }
      return got > 0 ? costSum / got : (it.purch || 0);
    };
    for (const e of events) {
      if (e.type === 'purch') {
        if (e.netQty > 1e-9) layers.push({ qty: e.netQty, rate: e.rate });
        else if (e.netQty < -1e-9) withdraw(-e.netQty);
      } else if (e.type === 'repair') {
        if (e.netQty < -1e-9) withdraw(-e.netQty);
      } else if (e.type === 'sale') {
        if (e.netQty < -1e-9) {
          const q = -e.netQty; const costRate = withdraw(q);
          if (e.date >= fromDate && e.date <= toDate) { revenue += q * (e.rate || 0); cost += q * costRate; qty += q; }
        } else if (e.netQty > 1e-9) {
          layers.push({ qty: e.netQty, rate: it.purch || 0 });
        }
      }
    }
  }
  return { revenue: Math.round(revenue), cost: Math.round(cost), profit: Math.round(revenue - cost), qty: Math.round(qty * 100) / 100 };
}

// Per-sale FIFO cost/margin detail for ONE item, keyed by sale id — same
// replay as getFifoSaleMarginMap above, but kept per-transaction instead of
// aggregated, for the Item Ledger's per-row margin column. A sales-return
// line resolves to isReturn:true with no margin (a return isn't a sale).
export function getFifoSaleCosts(itemId) {
  const it = state.items.find(x => x.id === itemId);
  if (!it) return {};
  const events = [];
  for (const p of state.purchases) for (const l of (p.items || [])) {
    if (l.itemId !== itemId) continue;
    events.push({ date: p.date, seq: p.createdAt || '', type: 'purch', netQty: l.qty || 0, rate: l.rate || 0 });
  }
  for (const s of state.sales) for (const x of (s.items || [])) {
    if (x.itemId !== itemId) continue;
    events.push({ date: s.date, seq: s.createdAt || '', type: 'sale', netQty: -(x.qty || 0), rate: x.rate, saleId: s.id });
  }
  for (const r of state.repairExpenses) {
    if (r.itemId !== itemId) continue;
    events.push({ date: r.date, seq: r.createdAt || '', type: 'repair', netQty: -(r.qty || 0) });
  }
  events.sort((a, b) => a.date.localeCompare(b.date) || a.seq.localeCompare(b.seq));
  const netFromEvents = events.reduce((s, e) => s + e.netQty, 0);
  const plug = (it.stock || 0) - netFromEvents;
  const layers = [];
  if (plug > 1e-9) layers.push({ qty: plug, rate: it.purch || 0 });
  const withdraw = qty => {
    let need = qty, costSum = 0, got = 0;
    while (need > 1e-9 && layers.length) {
      const l = layers[0]; const take = Math.min(l.qty, need);
      costSum += take * l.rate; got += take; l.qty -= take; need -= take;
      if (l.qty <= 1e-9) layers.shift();
    }
    if (need > 1e-9) { costSum += need * (it.purch || 0); got += need; }
    return got > 0 ? costSum / got : (it.purch || 0);
  };
  const results = {};
  for (const e of events) {
    if (e.type === 'purch') {
      if (e.netQty > 1e-9) layers.push({ qty: e.netQty, rate: e.rate });
      else if (e.netQty < -1e-9) withdraw(-e.netQty);
    } else if (e.type === 'repair') {
      if (e.netQty < -1e-9) withdraw(-e.netQty);
    } else if (e.type === 'sale') {
      if (e.netQty < -1e-9) {
        const qty = -e.netQty; const costRate = withdraw(qty);
        const marginPerPc = Math.round(((e.rate || 0) - costRate) * 100) / 100;
        const marginPct = e.rate > 0 ? Math.round(marginPerPc / e.rate * 1000) / 10 : 0;
        results[e.saleId] = { qty, rate: e.rate, costRate: Math.round(costRate * 100) / 100, marginPerPc, marginPct, isReturn: false };
      } else if (e.netQty > 1e-9) {
        layers.push({ qty: e.netQty, rate: it.purch || 0 });
        results[e.saleId] = { qty: -e.netQty, rate: e.rate, costRate: it.purch || 0, marginPerPc: 0, marginPct: 0, isReturn: true };
      }
    }
  }
  return results;
}

// "What are we actually selling this at lately?" — the mode (most frequent
// rate) among an item's N most recent in-app sale lines, tie-broken toward
// whichever of the tied rates was charged more recently. Returns null (not 0)
// when the item has no in-app sale history at all, so callers can tell
// "never sold in-app" apart from "sold for ₹0". Return lines (x.qty<0) are
// excluded — a return doesn't reflect a current selling rate.
export function getRecentModeSaleRate(itemId, n = 10) {
  if (!itemId) return null;
  const lines = [];
  for (const s of state.sales) {
    for (const x of (s.items || [])) {
      if (x.itemId === itemId && (x.qty || 0) > 0 && x.rate != null) {
        lines.push({ date: s.date, seq: s.createdAt || '', rate: x.rate });
      }
    }
  }
  if (!lines.length) return null;
  lines.sort((a, b) => b.date.localeCompare(a.date) || b.seq.localeCompare(a.seq));
  const recent = lines.slice(0, n);
  const freq = {};
  for (const l of recent) freq[l.rate] = (freq[l.rate] || 0) + 1;
  const maxCount = Math.max(...Object.values(freq));
  for (const l of recent) if (freq[l.rate] === maxCount) return l.rate; // first hit = most recent among the modes
  return recent[0].rate;
}

// Balance Sheet: Assets (cash/bank + inventory value + receivables) vs Liabilities
// (payables to suppliers/transporters). Equity is a computed residual (Assets minus
// Liabilities) — there's no separate owner's-capital ledger to draw from.
export function getBalanceSheet() {
  const cash = ACCOUNTS.reduce((s, a) => s + getAccBal(a.id), 0);
  const fifoMap = getFifoStockMap();
  const inventoryValue = state.items.filter(it => it.active !== false).reduce((s, it) => s + (fifoMap[it.id] ? fifoMap[it.id].value : ((it.stock || 0) * (it.purch || 0))), 0);
  const receivables = state.parties.filter(p => p.type === 'customer').reduce((s, p) => { const od = getCustomerBal(p.id); return s + (od > 0 ? od : 0); }, 0);
  const payables = state.parties.filter(p => p.type === 'supplier' || p.type === 'transporter').reduce((s, p) => { const bal = getSupplierBal(p.id); return s + (bal > 0 ? bal : 0); }, 0);
  const assets = cash + inventoryValue + receivables;
  const equity = assets - payables;
  return { cash, inventoryValue, receivables, assets, payables, equity };
}

export function getBestPurchRate(supplierName, itemId, def) {
  let best = null;
  for (const p of state.purchases) {
    if (p.partyName !== supplierName) continue;
    const line = p.items.find(l => l.itemId === itemId);
    if (line && (!best || p.date > best.date)) best = { date: p.date, rate: line.rate };
  }
  return best ? best.rate : (def || 0);
}

// Expected inventory: qty from purchases that have been booked (ordered) but
// not yet received — stock isn't incremented until receivedDate is set, so
// this is the "incoming" quantity per item that isn't in getAllItemMetrics'
// stock figure yet.
export function getExpectedStockMap() {
  const map = {};
  for (const p of state.purchases) {
    if (p.receivedDate) continue;
    for (const l of (p.items || [])) {
      if (!l.itemId) continue;
      map[l.itemId] = (map[l.itemId] || 0) + (l.qty || 0);
    }
  }
  return map;
}

// Local-time "YYYY-MM" for a Date — NOT d.toISOString().slice(0,7), which
// converts to UTC first. For any timezone ahead of UTC (IST included), the
// 1st of the month at local midnight is still the last day of the PREVIOUS
// month in UTC, so that used to silently shift every monthly bucket back by
// one — dropping the current month's data entirely (it had no matching
// bucket at all) and mislabeling every other month in the trend.
const monthKey = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');

// Last N months of Sales / Expenses / Profit, for the Business Intelligence
// trend chart. COGS uses the same current-rate proxy as getPnL's new-sales
// leg (item.purchRaw × qty) applied across every month including historical
// ones — there's no per-month historical cost breakdown to draw on, so this
// is an approximation, consistent with the proxy already used elsewhere.
// Expenses are business-tagged categories only, pulled from the dated
// EXPENSE_SUBLEDGER entries (which reconcile exactly to HISTORICAL_EXPENSES)
// plus new in-app txns.
export function getMonthlyBI(numMonths = 12) {
  const now = new Date();
  const months = [];
  for (let i = numMonths - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push(monthKey(d));
  }
  const idx = new Map(months.map((m, i) => [m, i]));
  const sales = new Array(months.length).fill(0);
  const cogs = new Array(months.length).fill(0);
  const expenses = new Array(months.length).fill(0);

  const itemCost = new Map();
  for (const it of state.items) {
    const rate = it.purchRaw ?? it.purch ?? 0;
    if (it.internal) itemCost.set(it.internal.toLowerCase(), rate);
    if (it.customer) itemCost.set((it.customer || '').toLowerCase(), rate);
  }

  for (const r of getAllItemRows()) {
    if (r.dir !== 'out') continue;
    const i = idx.get((r.d || '').slice(0, 7));
    if (i === undefined) continue;
    // Returns now carry a negative qty/amount (see getAllItemRows), so adding
    // them directly nets them out of that month's sales/cogs instead of
    // double-counting them as additional revenue.
    sales[i] += r.isReturn ? -Math.abs(r.amount||0) : Math.abs(r.amount || 0);
    cogs[i] += (r.isReturn ? -Math.abs(r.qty||0) : Math.abs(r.qty || 0)) * (itemCost.get((r.item || '').toLowerCase()) || 0);
  }

  const bizCatIds = [...TXN_CATEGORIES, ...state.customCategories]
    .filter(c => c.dir === 'out' && c.party === null && getCategoryBiz(c.id) === 'business')
    .map(c => c.id);
  for (const catId of bizCatIds) {
    const subs = EXPENSE_SUBLEDGER[catId] || {};
    for (const head in subs) {
      for (const [d, amt] of subs[head]) {
        const i = idx.get((d || '').slice(0, 7));
        if (i !== undefined) expenses[i] += amt;
      }
    }
  }
  for (const t of state.txns) {
    if (t.type !== 'payment' || !bizCatIds.includes(t.category)) continue;
    const i = idx.get((t.date || '').slice(0, 7));
    if (i !== undefined) expenses[i] += t.amount;
  }

  return months.map((m, i) => ({
    month: m,
    sales: Math.round(sales[i]),
    expenses: Math.round(expenses[i]),
    profit: Math.round(sales[i] - cogs[i] - expenses[i]),
  }));
}

// Same "last N months" trend as getMonthlyBI, but split into two lines of
// business — "Lighter" (item.cat === 'Lighter') vs everything else, folded
// into "Home Appliances". Sales & gross profit bucket cleanly since every
// sold item carries its own category. Receipts don't: a payment is recorded
// as one lump sum against a party, never against a specific invoice or item,
// so there's no stored fact about which category a given rupee of cash paid
// off. This walks each customer's full sale + payment history in date order
// — mirroring the FIFO oldest-bill-first matching getAllCustomerMetrics
// already uses for avg-payment-days — and as each payment consumes a bill,
// attributes the paid amount to that bill's category mix (proportionally,
// for a bill that mixes both categories). Returns are left out of the walk
// entirely (same simplification getAllCustomerMetrics makes): they're not
// cash received, and folding them in would mean guessing which of several
// open bills a return should close first. Any payment left over once the
// known bill queue is exhausted (an advance, or a balance carried in from
// before this customer's tracked history) is folded into Home Appliances,
// the dominant line of business — a small, documented approximation.
export function getMonthlyBIByCategory(numMonths = 12) {
  const now = new Date();
  const months = [];
  for (let i = numMonths - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push(monthKey(d));
  }
  const idx = new Map(months.map((m, i) => [m, i]));
  const mk = () => new Array(months.length).fill(0);
  const sales = { home: mk(), lighter: mk() };
  const cogs = { home: mk(), lighter: mk() };
  const receipts = { home: mk(), lighter: mk() };

  const nameToBucket = new Map();
  const itemCost = new Map();
  for (const it of state.items) {
    const bucket = it.cat === 'Lighter' ? 'lighter' : 'home';
    const rate = it.purchRaw ?? it.purch ?? 0;
    for (const n of [it.internal, it.name, it.customer, ...(it.mergedNames || [])]) {
      if (!n) continue;
      const k = n.toLowerCase().trim();
      nameToBucket.set(k, bucket);
      itemCost.set(k, rate);
    }
  }
  const bucketOf = name => nameToBucket.get(String(name || '').toLowerCase().trim()) || 'home';

  // Sales & gross profit — every out-bound item row, bucketed by its own name.
  for (const r of getAllItemRows()) {
    if (r.dir !== 'out') continue;
    const i = idx.get((r.d || '').slice(0, 7));
    if (i === undefined) continue;
    const bucket = bucketOf(r.item);
    const amt = r.isReturn ? -Math.abs(r.amount || 0) : Math.abs(r.amount || 0);
    const qty = r.isReturn ? -Math.abs(r.qty || 0) : Math.abs(r.qty || 0);
    sales[bucket][i] += amt;
    cogs[bucket][i] += qty * (itemCost.get(String(r.item || '').toLowerCase().trim()) || 0);
  }

  // A bill's items, split into home/lighter $ amounts scaled to sum exactly
  // to the bill's real total (so the category split never drifts from the
  // actual rupees a payment is clearing).
  const catBreakdown = (items, total, nameOf, qtyOf, rateOf) => {
    let home = 0, lighter = 0;
    for (const it of items) {
      const amt = (qtyOf(it) || 0) * (rateOf(it) || 0);
      if (bucketOf(nameOf(it)) === 'lighter') lighter += amt; else home += amt;
    }
    const sum = home + lighter;
    if (sum) { const s = total / sum; home *= s; lighter *= s; }
    else if (total) home = total;
    return { home, lighter };
  };

  for (const p of state.parties) {
    if (p.type !== 'customer') continue;
    const hist = PARTY_HISTORY[p.name] || {};
    const histBills = [];
    for (const [d, total, items] of (hist.s || [])) {
      for (const seg of splitPosNegItems(total, items)) {
        if (seg.isReturn) continue;
        histBills.push({ date: d, total: seg.total, cats: catBreakdown(seg.items || [], seg.total, x => x[0], x => x[1], x => x[2]) });
      }
    }
    const newBills = state.sales
      .filter(s => s.partyId === p.id && s.type !== 'return')
      .map(s => ({ date: s.date, total: s.total, cats: catBreakdown(s.items || [], s.total, x => x.customerName || x.itemName, x => x.qty, x => x.rate) }));
    const billQueue = [...histBills, ...newBills]
      .sort((a, b) => a.date.localeCompare(b.date))
      .map(b => ({ rem: b.total, cats: { ...b.cats } }));

    const payEvents = [
      ...(hist.p || []).map(([d, amt]) => ({ date: d, amt })),
      ...state.payments.filter(pm => pm.partyId === p.id).map(pm => ({ date: pm.date, amt: pm.amount })),
    ].sort((a, b) => a.date.localeCompare(b.date));

    let bi = 0;
    for (const pay of payEvents) {
      let remaining = pay.amt;
      const mi = idx.get((pay.date || '').slice(0, 7));
      while (remaining > 0.01 && bi < billQueue.length) {
        const bill = billQueue[bi];
        if (bill.rem <= 0.01) { bi++; continue; }
        const applied = Math.min(remaining, bill.rem);
        if (mi !== undefined) {
          const fracLighter = bill.rem > 0 ? bill.cats.lighter / bill.rem : 0;
          receipts.lighter[mi] += applied * fracLighter;
          receipts.home[mi] += applied * (1 - fracLighter);
        }
        const scale = (bill.rem - applied) / bill.rem;
        bill.cats.home *= scale; bill.cats.lighter *= scale;
        bill.rem -= applied; remaining -= applied;
        if (bill.rem <= 0.01) bi++;
      }
      if (remaining > 0.01 && mi !== undefined) receipts.home[mi] += remaining;
    }
  }

  return months.map((m, i) => ({
    month: m,
    home: { sales: Math.round(sales.home[i]), profit: Math.round(sales.home[i] - cogs.home[i]), receipts: Math.round(receipts.home[i]) },
    lighter: { sales: Math.round(sales.lighter[i]), profit: Math.round(sales.lighter[i] - cogs.lighter[i]), receipts: Math.round(receipts.lighter[i]) },
  }));
}

// Sales & gross profit for the last N months, bucketed by each item's real
// category (Gas Stove, Induction Cooker, Infrared, Lighter, ...) rather than
// just the Home Appliances/Lighter split above — a 12-month total per
// category, not a monthly trend (a 10-line trend chart isn't very readable).
// Receipts aren't included here: splitting them this finely would mean
// tracking a 10-way category mix per bill instead of 2 in the FIFO walk
// above, for a figure that's already an approximation with just 2 buckets.
export function getSalesProfitByCategory(numMonths = 12) {
  const now = new Date();
  const cutoff = monthKey(new Date(now.getFullYear(), now.getMonth() - (numMonths - 1), 1));

  const itemCost = new Map();
  const itemCat = new Map();
  for (const it of state.items) {
    const rate = it.purchRaw ?? it.purch ?? 0;
    const cat = it.cat || 'Other';
    for (const n of [it.internal, it.name, it.customer, ...(it.mergedNames || [])]) {
      if (!n) continue;
      const k = n.toLowerCase().trim();
      itemCost.set(k, rate);
      itemCat.set(k, cat);
    }
  }

  const byCat = new Map();
  for (const r of getAllItemRows()) {
    if (r.dir !== 'out') continue;
    if ((r.d || '').slice(0, 7) < cutoff) continue;
    const key = String(r.item || '').toLowerCase().trim();
    const cat = itemCat.get(key) || 'Other';
    const amt = r.isReturn ? -Math.abs(r.amount || 0) : Math.abs(r.amount || 0);
    const qty = r.isReturn ? -Math.abs(r.qty || 0) : Math.abs(r.qty || 0);
    const cur = byCat.get(cat) || { sales: 0, cogs: 0, qty: 0 };
    cur.sales += amt;
    cur.cogs += qty * (itemCost.get(key) || 0);
    cur.qty += qty;
    byCat.set(cat, cur);
  }

  return [...byCat.entries()].map(([cat, v]) => ({
    cat, sales: Math.round(v.sales), profit: Math.round(v.sales - v.cogs), qty: Math.round(v.qty),
    margin: v.sales ? Math.round((v.sales - v.cogs) / v.sales * 100) : 0,
  })).sort((a, b) => b.sales - a.sales);
}

// A customer-facing name correction made in the invoice preview applies to
// every customer's invoices for that item, but is not written back to the
// shared item master (item.customer is the "system" name used by search).
export function getItemCustomerName(itemId, fallback) {
  return state.itemCustomerNames?.[itemId] || fallback;
}
export function setItemCustomerName(itemId, name) {
  if (!itemId || !name) return;
  state.itemCustomerNames = { ...state.itemCustomerNames, [itemId]: name };
  save.itemCustomerNames();
}

export function getPrevItems(partyName) {
  if (!partyName) return [];
  const seen = new Set(); const result = [];
  for (const [pn,iName,rate] of LAST_RATES) {
    if (pn !== partyName || seen.has(iName)) continue;
    seen.add(iName);
    const item = state.items.find(it =>
      (it.internal||'').toLowerCase()===iName.toLowerCase() ||
      (it.customer||'').toLowerCase()===iName.toLowerCase()
    ) || { id:'_'+iName, internal:iName, customer:iName, cat:'Other', peti:getPetiFromName(iName), sale:rate, purch:0 };
    if (item.active === false) continue;
    result.push({ item, rate });
  }
  return result.sort((a,b) => a.item.cat.localeCompare(b.item.cat)||(a.item.customer||a.item.internal).localeCompare(b.item.customer||b.item.internal));
}

// ── GLOBAL STATE ───────────────────────────────────────────────────────────────
export const state = {
  items:            [...SEED_ITEMS],
  parties:          [...SEED_PARTIES],
  sales:            [],
  payments:         [],
  purchases:        [],
  supplierPayments: [],
  freightCharges:   [],
  txns:             [],
  followups:        [],
  templates:        {},
  phoneEdits:       {},
  partyTransport:   {},
  cityTransport:    {},
  customCategories: [],
  categoryTags:     {},
  orders:           [],
  dispatchPhone:    '',
  cityOrder:        [],
  itemCustomerNames: {},
  repairExpenses:   [],
  stockAdjustments: [],
  // Item ids temporarily dismissed from the Dashboard's "Out of Stock" list
  // (item #4 of the dashboard request) — a soft hide, not a data change:
  // the item is still genuinely out of stock, the user just doesn't want it
  // cluttering the board right now (e.g. a discontinued line they're not
  // reordering). Toggled back on from the same card's "Hidden (n)" sub-view.
  dashboardHiddenOOS: [],
  // Owner/Staff role mapping, keyed by signed-in Supabase auth email — see
  // computeRole()/useRole() in App.jsx. NOT a persisted-at-boot "current
  // user" field: currentUserEmail below is runtime-only, set fresh from
  // DB.getSession() each load, never saved/synced itself.
  users:            [],
  currentUserEmail: '',
  loaded:           false,
};

export const save = {
  sales:            () => DB.set('data:sales',            state.sales),
  payments:         () => DB.set('data:payments',          state.payments),
  purchases:        () => DB.set('data:purchases',         state.purchases),
  supplierPayments: () => DB.set('data:supplierPayments',  state.supplierPayments),
  freightCharges:   () => DB.set('data:freightCharges',    state.freightCharges),
  txns:             () => DB.set('data:txns',              state.txns),
  followups:        () => DB.set('data:followups',         state.followups),
  templates:        () => DB.set('data:templates',         state.templates),
  phoneEdits:       () => DB.set('data:phoneEdits',        state.phoneEdits),
  partyTransport:   () => DB.set('data:partyTransport',    state.partyTransport),
  cityTransport:    () => DB.set('data:cityTransport',     state.cityTransport),
  parties:          () => DB.set('data:parties',           state.parties),
  items:            () => DB.set('data:items',             state.items),
  categoryTags:     () => DB.set('data:categoryTags',      state.categoryTags),
  customCategories: () => DB.set('data:customCategories',  state.customCategories),
  orders:           () => DB.set('data:orders',            state.orders),
  dispatchPhone:    () => DB.set('data:dispatchPhone',     state.dispatchPhone),
  itemCustomerNames: () => DB.set('data:itemCustomerNames', state.itemCustomerNames),
  cityOrder:        () => DB.set('data:cityOrder',         state.cityOrder),
  repairExpenses:   () => DB.set('data:repairExpenses',    state.repairExpenses),
  stockAdjustments: () => DB.set('data:stockAdjustments',  state.stockAdjustments),
  users:            () => DB.set('data:users',             state.users),
  dashboardHiddenOOS: () => DB.set('data:dashboardHiddenOOS', state.dashboardHiddenOOS),
};

// ── Dashboard: Out of Stock hide/unhide (request #4) ────────────────────────
export function hideFromDashboardOOS(itemId) {
  if (state.dashboardHiddenOOS.includes(itemId)) return;
  state.dashboardHiddenOOS = [...state.dashboardHiddenOOS, itemId];
  save.dashboardHiddenOOS();
}
export function unhideFromDashboardOOS(itemId) {
  state.dashboardHiddenOOS = state.dashboardHiddenOOS.filter(id => id !== itemId);
  save.dashboardHiddenOOS();
}

// ── Reorder levels (request #5/#6) ───────────────────────────────────────────
// A per-item `reorderLevel` (number of pcs) — undefined/0 means "no level
// set, don't flag this item". Stored directly on the item record (same
// pattern as every other per-item field) rather than a separate map, so it
// travels with the item through the existing items sync/save path with no
// new top-level state key needed.
export function setReorderLevel(itemId, level) {
  const lvl = Math.max(0, Number(level) || 0);
  state.items = state.items.map(it => it.id === itemId ? { ...it, reorderLevel: lvl } : it);
  save.items();
}
// Items at or below their set reorder level (and still active) — the
// Dashboard's "Reorder" list. Out-of-stock items (stock<=0) already have
// their own dedicated board section, so this list is for the "running low,
// not yet zero" warning zone; an item sitting at 0 with a reorder level set
// still qualifies here too (it needs reordering more than anyone), it just
// also shows up in the Out of Stock card.
export function getReorderItems() {
  return state.items.filter(it => it.active !== false && (it.reorderLevel || 0) > 0 && (it.stock || 0) <= it.reorderLevel)
    .sort((a, b) => (a.stock || 0) - (b.stock || 0));
}

// Learns which transporter is typically used for each delivery city, from
// the "Transport" field entered on past sales — so New Sale can suggest a
// sensible default once a few sales to that city exist, while still letting
// the user pick a different one from the list or type a new name.
export function recordCityTransport(city, transportName) {
  const name = (transportName || '').trim();
  if (!city || !name) return;
  const forCity = { ...(state.cityTransport[city] || {}) };
  forCity[name] = (forCity[name] || 0) + 1;
  state.cityTransport = { ...state.cityTransport, [city]: forCity };
  save.cityTransport();
}

export function getTopTransportersForCity(city, limit = 5) {
  const forCity = state.cityTransport[city] || {};
  return Object.entries(forCity).sort((a, b) => b[1] - a[1]).slice(0, limit).map(([name]) => name);
}

export function getAllTransporterNames() {
  const counts = {};
  for (const s of state.sales) {
    const t = (s.transport || '').trim();
    if (t) counts[t] = (counts[t] || 0) + 1;
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([name]) => name);
}

// ── LIVE UPDATE (cross-tab) ───────────────────────────────────────────────────
// state is a plain in-memory object with no reactivity of its own — components
// re-render themselves after their own save() calls, but a change made in a
// different browser tab (opened via the sidebar's "open in new tab" links)
// has no way to reach them otherwise. Since DB.js sits directly on
// localStorage, the native `storage` event fires in every OTHER tab whenever
// one tab writes a key — so we use it to pull that slice of state back in
// and notify subscribers, without touching same-tab behavior at all.
const STORAGE_KEY_TO_STATE = {
  'data:sales':'sales', 'data:payments':'payments', 'data:purchases':'purchases',
  'data:supplierPayments':'supplierPayments', 'data:freightCharges':'freightCharges',
  'data:txns':'txns', 'data:followups':'followups', 'data:templates':'templates',
  'data:phoneEdits':'phoneEdits', 'data:partyTransport':'partyTransport',
  'data:cityTransport':'cityTransport', 'data:parties':'parties', 'data:items':'items',
  'data:categoryTags':'categoryTags', 'data:customCategories':'customCategories',
  'data:orders':'orders', 'data:dispatchPhone':'dispatchPhone', 'data:itemCustomerNames':'itemCustomerNames',
  'data:cityOrder':'cityOrder', 'data:repairExpenses':'repairExpenses',
  'data:stockAdjustments':'stockAdjustments', 'data:users':'users',
  'data:dashboardHiddenOOS':'dashboardHiddenOOS',
};
const liveListeners = new Set();
export function subscribeLive(fn) { liveListeners.add(fn); return () => liveListeners.delete(fn); }
if (typeof window !== 'undefined') {
  window.addEventListener('storage', (e) => {
    const key = STORAGE_KEY_TO_STATE[e.key];
    if (!key || e.newValue == null) return;
    try { state[key] = JSON.parse(e.newValue); } catch { return; }
    liveListeners.forEach(fn => fn());
  });
}

// Same live-refresh path as the same-device `storage` listener above, but
// fed by db.js's Supabase Realtime subscription instead — this is how a
// change made on a DIFFERENT device reaches this tab's already-rendered
// components, reusing the exact same STORAGE_KEY_TO_STATE map and notify
// mechanism (components already subscribed via subscribeLive() need no
// changes at all to pick up a remote-origin update).
DB.onRemoteUpdate((storageKey, value) => {
  const key = STORAGE_KEY_TO_STATE[storageKey];
  if (!key) return;
  state[key] = value;
  liveListeners.forEach(fn => fn());
});

export function updatePartyPhone(partyId, phone) {
  state.phoneEdits[partyId] = phone;
  state.parties = state.parties.map(p => p.id===partyId ? {...p,phone} : p);
  save.phoneEdits();
  save.parties();
}

export async function loadAll() {
  if (state.loaded) return;
  try {
    const [sales,payments,pe,pt,fu,txns,parties,items,purchases,supplierPayments,freightCharges,customCategories,categoryTags,cityTransport,orders,dispatchPhone,itemCustomerNames,cityOrder,repairExpenses,stockAdjustments,users,dashboardHiddenOOS] = await Promise.all([
      DB.get('data:sales'), DB.get('data:payments'), DB.get('data:phoneEdits'),
      DB.get('data:partyTransport'), DB.get('data:followups'), DB.get('data:txns'),
      DB.get('data:parties'), DB.get('data:items'), DB.get('data:purchases'), DB.get('data:supplierPayments'),
      DB.get('data:freightCharges'), DB.get('data:customCategories'), DB.get('data:categoryTags'),
      DB.get('data:cityTransport'), DB.get('data:orders'), DB.get('data:dispatchPhone'), DB.get('data:itemCustomerNames'),
      DB.get('data:cityOrder'), DB.get('data:repairExpenses'), DB.get('data:stockAdjustments'), DB.get('data:users'),
      DB.get('data:dashboardHiddenOOS'),
    ]);
    if (sales)    state.sales    = sales;
    if (payments) state.payments = payments;
    if (txns)     state.txns     = txns;
    if (fu)       state.followups = fu;
    if (parties)  state.parties  = parties;
    if (items)    state.items    = items;
    if (purchases)        state.purchases        = purchases;
    if (supplierPayments) state.supplierPayments  = supplierPayments;
    if (freightCharges)   state.freightCharges    = freightCharges;
    if (customCategories) state.customCategories  = customCategories;
    if (categoryTags)     state.categoryTags      = categoryTags;
    if (orders)           state.orders            = orders;
    if (dispatchPhone)    state.dispatchPhone      = dispatchPhone;
    if (itemCustomerNames) state.itemCustomerNames = itemCustomerNames;
    if (cityOrder)        state.cityOrder         = cityOrder;
    if (repairExpenses)   state.repairExpenses    = repairExpenses;
    if (stockAdjustments) state.stockAdjustments  = stockAdjustments;
    if (users)            state.users             = users;
    if (dashboardHiddenOOS) state.dashboardHiddenOOS = dashboardHiddenOOS;
    // One-time migration: persisted items predate two corrections in the seed
    // data — (1) `purch` used to be a flat 22% of `sale` instead of a real
    // rate, hiding per-item margins, and (2) `stock`/`stockValue` weren't yet
    // sourced from the Purchase sheet's authoritative Closing Stock (qty) and
    // Closing Stock Value columns. Refresh master-data fields from the current
    // seed by id. Stock/stockValue only refresh for items with no real
    // in-app sale/purchase against them yet, so live inventory the user has
    // actually tracked here isn't clobbered. Runs after state.sales/state.purchases
    // are populated above, so the "already touched" check sees real data.
    // Also adds any items new to the seed since.
    if (items) {
      const migrated = await DB.get('data:itemsMigratedV9');
      if (!migrated) {
        const seedById = new Map(SEED_ITEMS.map(s => [s.id, s]));
        const touchedIds = new Set([
          ...state.sales.flatMap(s => (s.items || []).map(l => l.itemId)),
          ...state.purchases.flatMap(p => (p.items || []).map(l => l.itemId)),
        ]);
        state.items = state.items.map(it => {
          const seed = seedById.get(it.id);
          if (!seed) return it;
          const refreshed = { ...it, internal: seed.internal, customer: seed.customer, supplier: seed.supplier, cat: seed.cat, peti: seed.peti, sale: seed.sale, purch: seed.purch, purchRaw: seed.purchRaw, sortOrder: seed.sortOrder, subHead: seed.subHead, lastPurchDate: seed.lastPurchDate };
          if (!touchedIds.has(it.id)) {
            refreshed.stock = seed.stock;
            if (seed.stockValue !== undefined) refreshed.stockValue = seed.stockValue;
          }
          return refreshed;
        });
        const existingIds = new Set(state.items.map(i => i.id));
        state.items = [...state.items, ...SEED_ITEMS.filter(s => !existingIds.has(s.id))];
        save.items();
        await DB.set('data:itemsMigratedV9', true);
      }
    }
    // One-time, narrowly-scoped text fix: the product line sold as "Indus
    // Surya9" was renamed to "Unidus Surya9" (seed data already carries the
    // new name), but a browser that already ran the itemsMigratedV9 refresh
    // above has the old spelling baked into real, already-entered data —
    // the items master AND every historical sale/purchase line that copied
    // the item's name at entry time (itemName/customerName/supplierName —
    // see getAllItemRows below for why those copies exist separately from
    // the item master). Doing this as a plain string substitution rather
    // than another full seed-refresh (like V9 above) means it can't clobber
    // anything else the user has since edited by hand (rates, subHead
    // reassignments via Sub-Categories, etc.) — it only touches the literal
    // word "Indus".
    {
      const fixIndusText = s => {
        if (typeof s !== 'string' || !s || !/\bindus\b/i.test(s)) return s;
        return s.replace(/\bindus\b/gi, m => m === m.toUpperCase() ? 'UNIDUS' : (m[0] === m[0].toUpperCase() ? 'Unidus' : 'unidus'));
      };
      const hasIndus = s => typeof s === 'string' && /\bindus\b/i.test(s);
      const unidusFixed = await DB.get('data:unidusRenameV1');
      if (!unidusFixed) {
        let changedItems = false, changedSales = false, changedPurchases = false;
        state.items = state.items.map(it => {
          let touched = false;
          const next = { ...it };
          for (const f of ['internal', 'customer', 'supplier', 'subHead', 'dispatchName']) {
            if (hasIndus(next[f])) { next[f] = fixIndusText(next[f]); touched = true; }
          }
          if (touched) changedItems = true;
          return touched ? next : it;
        });
        state.sales = state.sales.map(s => {
          if (!(s.items || []).some(x => hasIndus(x.itemName) || hasIndus(x.customerName))) return s;
          changedSales = true;
          return { ...s, items: s.items.map(x => ({ ...x, itemName: fixIndusText(x.itemName), customerName: fixIndusText(x.customerName) })) };
        });
        state.purchases = state.purchases.map(p => {
          if (!(p.items || []).some(x => hasIndus(x.itemName) || hasIndus(x.customerName) || hasIndus(x.supplierName))) return p;
          changedPurchases = true;
          return { ...p, items: p.items.map(x => ({ ...x, itemName: fixIndusText(x.itemName), customerName: fixIndusText(x.customerName), supplierName: fixIndusText(x.supplierName) })) };
        });
        if (changedItems) save.items();
        if (changedSales) save.sales();
        if (changedPurchases) save.purchases();
        await DB.set('data:unidusRenameV1', true);
      }
    }
    // One-time, generic fix for sub-category (subHead) groups that have
    // silently forked into near-duplicates: e.g. some "1B Variants" items
    // and one stray "1b variants " item (extra space) or "1B variants"
    // (different case) — same intended group, different exact spelling.
    // The main Item Masters table and mobile Inventory list only print a
    // group header when adjacent items' subHead text differs, and they
    // only end up adjacent via sortOrder, which setItemSubHead (App.jsx)
    // manages whenever someone reassigns a subHead by hand — but that
    // matching used to be an exact string compare, so a slightly
    // differently-spelled entry (typed once, long before that function
    // existed, or added outside it) would sit elsewhere in sortOrder and
    // show up as its own stray one-item group, exactly looking "misfiled"
    // or "missing" from its real group even though the data was always
    // basically right. This migration finds such normalized-duplicate
    // groups (same cat + same subHead once trimmed/case-folded/whitespace-
    // collapsed, but not byte-for-byte identical), picks whichever exact
    // spelling the majority already uses as canonical, rewrites the
    // stragglers' subHead to that spelling, and moves them to sit right
    // after the rest of their real group in sortOrder — i.e. exactly what
    // manually reassigning them via Sub-Categories would now do (see the
    // normalized matching added to setItemSubHead), just automatically and
    // once. It deliberately does NOT touch items whose subHead is a
    // genuinely different group name (that's a real reclassification
    // decision, left to Sub-Categories) — it only merges near-duplicate
    // spellings of what's already, unambiguously, the same group.
    {
      const subHeadFixed = await DB.get('data:subHeadDedupeV1');
      if (!subHeadFixed) {
        const norm = s => (s || '').trim().toLowerCase().replace(/\s+/g, ' ');
        const groups = new Map(); // "catnormSubHead" -> items[]
        for (const it of state.items) {
          if (!it.subHead) continue;
          const key = (it.cat || '') + '' + norm(it.subHead);
          if (!groups.has(key)) groups.set(key, []);
          groups.get(key).push(it);
        }
        const rewrites = new Map(); // id -> { subHead, sortOrder }
        for (const members of groups.values()) {
          const spellings = new Set(members.map(m => m.subHead));
          if (spellings.size <= 1) continue; // already consistent, nothing to merge
          const counts = new Map();
          for (const m of members) counts.set(m.subHead, (counts.get(m.subHead) || 0) + 1);
          let canonical = members[0].subHead, bestCount = -1, bestSort = Infinity;
          for (const m of members) {
            const c = counts.get(m.subHead);
            const s = m.sortOrder ?? Infinity;
            if (c > bestCount || (c === bestCount && s < bestSort)) { bestCount = c; bestSort = s; canonical = m.subHead; }
          }
          const canonicalMembers = members.filter(m => m.subHead === canonical);
          let anchor = Math.max(...canonicalMembers.map(m => m.sortOrder ?? 0));
          for (const m of members) {
            if (m.subHead === canonical) continue;
            anchor += 0.5;
            rewrites.set(m.id, { subHead: canonical, sortOrder: anchor });
          }
        }
        if (rewrites.size) {
          state.items = state.items.map(it => rewrites.has(it.id) ? { ...it, ...rewrites.get(it.id) } : it);
          save.items();
        }
        await DB.set('data:subHeadDedupeV1', true);
      }
    }
    // One-time correction: a ₹100,000 payment dated 2026-06-29 was recorded
    // against OCB Global (P0412) but actually belongs to VK Industries
    // (P0413) on 2026-06-27 — supplierHistory.js was corrected to reflect
    // this. Supplier `ob` is a frozen net balance computed once from history
    // rather than derived live from it, so moving the ledger entry alone
    // doesn't shift a balance already persisted to a browser from before this
    // fix — nudge both suppliers' ob by the same ₹100,000 the one time this
    // runs, composing on top of whatever ob is currently persisted.
    if (parties) {
      const fixed = await DB.get('data:fixOcbVkJun29Payment');
      if (!fixed) {
        state.parties = state.parties.map(p => {
          if (p.id === 'P0412') return { ...p, ob: (p.ob||0) + 100000 };
          if (p.id === 'P0413') return { ...p, ob: (p.ob||0) - 100000 };
          return p;
        });
        save.parties();
        await DB.set('data:fixOcbVkJun29Payment', true);
      }
    }
    if (pe) {
      state.phoneEdits = pe;
      state.parties = state.parties.map(p => state.phoneEdits[p.id]!==undefined ? {...p,phone:state.phoneEdits[p.id]} : p);
    }
    if (pt) state.partyTransport = pt;
    if (cityTransport) state.cityTransport = cityTransport;
    const tpls = await DB.get('data:templates');
    if (tpls) state.templates = {...DEFAULT_TEMPLATES, ...tpls};
    else state.templates = {...DEFAULT_TEMPLATES};
  } catch(e) { console.error('loadAll error', e); }
  state.loaded = true;
}

// ── TEMPLATES ─────────────────────────────────────────────────────────────────
export const DEFAULT_TEMPLATES = {
  invoice: `📅 {date}  |  Estimate

*Party:* {partyName}
{transport}
{freight}

*ITEMS*
{items}

*TOTAL: {total}*
{credit}`,
  payment: `Namaste Bhaiya 🙏
Apka *{amount}* outstanding balance hai.
Kripya payment ka dekh lijiye.

*{partyName}*
{invoiceDetails}
— Pratik, Upkaran`,
  followup: `Namaste Bhaiya 🙏
*{partyName}* — aapki last purchase *{lastItem}* thi *{lastDate}* ko.
Kya ab naya order dena hai? Stock available hai.
— Pratik, Upkaran`,
  delivery: `Namaste 🙏
*{partyName}*, aapka order deliver ho gaya hai:
{items}
{lr}
Kripya *confirm* kar dijiye ki saaman sahi salamat prapt ho gaya hai 🙏
— Pratik, Upkaran`,
  dispatch: `*DISPATCH ORDER*
📅 {date}

*Party:* {partyName}
*City:* {city}

*ITEMS*
{items}

*Total Qty:* {totalQty} pcs
*Total Amount:* {totalAmount}
{transport}`,
};

export function applyTemplate(tpl, vars) {
  let s = (state.templates[tpl] || DEFAULT_TEMPLATES[tpl] || '');
  Object.entries(vars).forEach(([k,v]) => { s = s.replaceAll('{'+k+'}', v||''); });
  return s.replace(/\n{3,}/g, '\n\n').trim();
}

export async function loadTemplates() {
  try {
    const t = await DB.get('data:templates');
    if (t) state.templates = { ...DEFAULT_TEMPLATES, ...t };
    else    state.templates = { ...DEFAULT_TEMPLATES };
  } catch { state.templates = { ...DEFAULT_TEMPLATES }; }
}
