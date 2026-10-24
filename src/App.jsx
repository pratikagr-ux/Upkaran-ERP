// ─────────────────────────────────────────────────────────────────────────────
// Upkaran ERP v2.0 — Main application
// All components in one file for Claude Code to extend/split further.
// Data is imported from src/data/ — no more baked-in seed data.
// Storage uses localStorage via src/db.js (same API as artifact window.storage).
// ─────────────────────────────────────────────────────────────────────────────

import { Fragment, useState, useEffect, useMemo, useCallback } from 'react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { C, fmt, fmtL, fmtD, today, uid, ROUTES, ACCOUNTS, ACC_MAP, TXN_CATEGORIES, HISTORICAL_EXPENSES, TAG_COLORS, APP_VERSION } from './constants.js';
import { DB } from './db.js';
import { supabase } from './supabaseClient.js';
import { state, save, loadAll, updatePartyPhone, getBestRate, gPS, getPrevItems, parseQty, parseResponse, getAccBal, getSupplierBal, getCustomerBal, getPartyBills, getSupplierBills, getBestPurchRate, getAllCustomerMetrics, getAllVendorMetrics, getExpenseReport, setCategoryBiz, getCategorySubLedger, getAllItemRows, getItemMetrics, getAllItemMetrics, getExpectedStockMap, getFifoStockMap, getRecentModeSaleRate, getMonthlyBI, getMonthlyBIByCategory, getSalesProfitByCategory, getPnL, getBalanceSheet, getRepairExpenseTotal, splitPosNegItems, recordCityTransport, getTopTransportersForCity, getAllTransporterNames, subscribeLive, getItemCustomerName, setItemCustomerName, PARTY_INVOICES, PARTY_HISTORY, SUPPLIER_HISTORY, ACCOUNT_OB, ACCOUNT_HISTORY, applyTemplate, DEFAULT_TEMPLATES } from './store.js';

// ── CONVENIENT ALIASES (same pattern as monolith) ─────────────────────────────
const _p  = () => state.parties;
const _i  = () => state.items;
const _s  = () => state.sales;
const _py = () => state.payments;
const _t  = () => state.txns;
const _fu = () => state.followups;

// delta>0 = stock in (purchase/return), delta<0 = stock out (sale)
function adjustStock(itemId, delta) {
  if (!itemId || !delta) return;
  const it = state.items.find(x => x.id === itemId);
  if (!it) return;
  state.items = state.items.map(x => x.id === itemId ? { ...x, stock: (x.stock||0) + delta } : x);
  save.items();
}

// Folds a duplicate party record into another — every in-app sale, payment,
// and follow-up gets re-pointed to the target's id, the opening balances add
// up, a blank phone is backfilled from the source, and the source's name is
// remembered on the target (mergedNames) so getPartyBills keeps pulling its
// historical invoices in too. The source record is then dropped from
// state.parties entirely — with everything re-pointed first, nothing is left
// referencing it, and removing it (rather than just hiding it) means every
// existing party picker/list across the app already excludes it for free.
function mergeParties(sourceId, targetId) {
  if (!sourceId || !targetId || sourceId === targetId) return;
  const source = state.parties.find(p => p.id === sourceId);
  const target = state.parties.find(p => p.id === targetId);
  if (!source || !target) return;

  state.sales = state.sales.map(s => s.partyId === sourceId ? { ...s, partyId: targetId, partyName: target.name } : s);
  save.sales();
  state.payments = state.payments.map(p => p.partyId === sourceId ? { ...p, partyId: targetId, partyName: target.name } : p);
  save.payments();
  state.followups = state.followups.map(f => f.partyId === sourceId ? { ...f, partyId: targetId } : f);
  save.followups();

  state.parties = state.parties
    .map(p => p.id === targetId ? {
      ...p,
      ob: (p.ob||0) + (source.ob||0),
      phone: p.phone || source.phone || '',
      mergedNames: [...new Set([...(p.mergedNames||[]), source.name, ...(source.mergedNames||[])])],
    } : p)
    .filter(p => p.id !== sourceId);
  save.parties();
}

// Folds a duplicate item into another — same pattern as mergeParties. Every
// sale/purchase line referencing the source itemId gets re-pointed, the
// stock/damaged/repaired counts add up, and the source's names are
// remembered on the target (mergedNames) so getItemMetrics keeps attributing
// historical (pre-app, name-matched) sales/purchase rows filed under the old
// name. The source item is then dropped from state.items entirely, so it
// disappears from every item picker across the app without needing to touch
// each one individually.
function mergeItems(sourceId, targetId) {
  if (!sourceId || !targetId || sourceId === targetId) return;
  const source = state.items.find(x => x.id === sourceId);
  const target = state.items.find(x => x.id === targetId);
  if (!source || !target) return;

  state.sales = state.sales.map(s => ({ ...s, items: (s.items||[]).map(l => l.itemId === sourceId ? { ...l, itemId: targetId } : l) }));
  save.sales();
  state.purchases = state.purchases.map(p => ({ ...p, items: (p.items||[]).map(l => l.itemId === sourceId ? { ...l, itemId: targetId } : l) }));
  save.purchases();

  const sourceNames = [source.internal, source.name, source.customer].filter(Boolean);
  state.items = state.items
    .map(x => x.id === targetId ? {
      ...x,
      stock: (x.stock||0) + (source.stock||0),
      damaged: (x.damaged||0) + (source.damaged||0),
      repaired: (x.repaired||0) + (source.repaired||0),
      mergedNames: [...new Set([...(x.mergedNames||[]), ...sourceNames, ...(source.mergedNames||[])])],
    } : x)
    .filter(x => x.id !== sourceId);
  save.items();
}

// Relocates a quantity of sellable stock from one item to another, unlike
// mergeItems: neither item is touched otherwise — no history reassignment, no
// mergedNames, and the source item is never deleted. For cases where two
// items are genuinely distinct SKUs and a purchase/count was simply logged
// under the wrong one, so only the stock number itself needs to move.
function moveItemStock(sourceId, targetId, qty) {
  if (!sourceId || !targetId || sourceId === targetId || !qty || qty <= 0) return;
  const source = state.items.find(x => x.id === sourceId);
  const target = state.items.find(x => x.id === targetId);
  if (!source || !target) return;
  const move = Math.min(qty, source.stock||0);
  if (move <= 0) return;
  state.items = state.items.map(x => {
    if (x.id === sourceId) return { ...x, stock: (x.stock||0) - move };
    if (x.id === targetId) return { ...x, stock: (x.stock||0) + move };
    return x;
  });
  save.items();
}

// Infrared/Induction goods coming back via a sales return aren't known-good —
// they land in the damaged bucket first and only become sellable stock again
// once someone repairs them (see repairStock below). Purchases, by contrast,
// go straight to sellable stock — the user pulls units out to "damaged"
// manually (moveToDamaged) once a problem is actually found.
const DAMAGE_TRACKED_CATS = new Set(['Infrared','Induction Cooker']);
const WARRANTY_OPTIONS = ['1yr warranty','1yr glass warranty','6 months warranty','no warranty'];

// A Jaali line added alongside an Infrared line (see the "Include Jaali"
// bundling in SaleEntry) is kept as its own line for stock/rate purposes, but
// customer/dispatch-facing text shows it folded into the Infrared line above
// it — one combined "Infrared +Jaali" bullet with summed rate and amount,
// rather than two separate bullets.
function mergeBundledLines(items){
  const out=[];
  for(const l of (items||[])){
    if(l.bundledJaali && out.length>0){
      const prev=out[out.length-1];
      prev.rate=(prev.rate||0)+(l.rate||0);
      prev.amount=(prev.amount||0)+(l.amount??(l.qty*l.rate));
    }else{
      out.push({...l});
    }
  }
  return out;
}
// Historical (pre-app) sale lines only ever recorded an item name, never an
// itemId — resolves one back to today's Item Master the same way
// getItemMetrics matches historical rows, so a return against an old
// invoice can still find the right item to credit.
function findItemByName(name) {
  const n = (name||'').toLowerCase().trim();
  if (!n) return null;
  return state.items.find(it => [it.internal, it.name, it.customer, ...(it.mergedNames||[])].filter(Boolean).some(s => s.toLowerCase().trim() === n)) || null;
}

// Credits a sales-return quantity to whichever bucket the user picked for
// that line (stock / damaged / repaired) — SalesReturnModal defaults
// damage-tracked categories to 'damaged' and everything else to 'stock',
// but lets the user override per line instead of silently assuming the
// goods came back faulty.
function applyReturnStock(itemId, qty, destination) {
  if (!itemId || !qty) return;
  const it = state.items.find(x => x.id === itemId);
  if (!it) return;
  const field = destination === 'repaired' ? 'repaired' : destination === 'damaged' ? 'damaged' : 'stock';
  state.items = state.items.map(x => x.id === itemId ? { ...x, [field]: (x[field]||0) + qty } : x);
  save.items();
}

// Logs a warranty/repair cost against a party — either a standalone entry
// (a faulty PCB comes back, 'return') or a reimbursement recorded alongside
// a payment receipt ('payment', see AccountsModule/OutstandingDetail). Kept
// as its own array (not folded into txns/payments) so it can be reported as
// a distinct P&L line and totalled per party regardless of which of the two
// triggered it.
function recordRepairExpense({partyId,partyName,date,itemId,itemName,qty,billedAmount,expenseAmount,source,linkedTxnId,notes}) {
  if (!partyId || !expenseAmount) return;
  const rec = {id:uid('RPX'),partyId,partyName,date,itemId:itemId||'',itemName:itemName||'',qty:qty||0,
    billedAmount:billedAmount||0,expenseAmount,source,linkedTxnId:linkedTxnId||null,notes:notes||'',createdAt:new Date().toISOString()};
  state.repairExpenses=[rec,...state.repairExpenses];
  save.repairExpenses();
}

// Moves qty from damaged into its OWN separate "repaired" pool — repaired
// stock is not folded back into plain sellable stock, it's tracked and (in
// future) sold on its own, typically at a discount, so fresh and repaired
// units never get mixed together in the same count.
function repairStock(itemId, qty) {
  if (!itemId || !qty || qty <= 0) return;
  const it = state.items.find(x => x.id === itemId);
  if (!it) return;
  const move = Math.min(qty, it.damaged||0);
  if (move <= 0) return;
  state.items = state.items.map(x => x.id === itemId ? { ...x, damaged: (x.damaged||0)-move, repaired: (x.repaired||0)+move } : x);
  save.items();
}

// The reverse of repairStock's damaged side — manually pull qty out of
// sellable stock into the damaged bucket, for goods found faulty after they
// were already counted as available (purchases land straight in stock, so
// this is how units get flagged as damaged going forward).
function moveToDamaged(itemId, qty) {
  if (!itemId || !qty || qty <= 0) return;
  const it = state.items.find(x => x.id === itemId);
  if (!it) return;
  const move = Math.min(qty, it.stock||0);
  if (move <= 0) return;
  state.items = state.items.map(x => x.id === itemId ? { ...x, stock: (x.stock||0)-move, damaged: (x.damaged||0)+move } : x);
  save.items();
}

// A correction, not a business transaction: for units that were never
// actually damaged — e.g. a purchase mis-routed into the damaged bucket by a
// since-fixed bug — this moves them straight back to plain stock, unlike
// repairStock, which is for genuinely damaged goods and feeds the separate
// repaired pool instead.
function correctDamagedToStock(itemId, qty) {
  if (!itemId || !qty || qty <= 0) return;
  const it = state.items.find(x => x.id === itemId);
  if (!it) return;
  const move = Math.min(qty, it.damaged||0);
  if (move <= 0) return;
  state.items = state.items.map(x => x.id === itemId ? { ...x, damaged: (x.damaged||0)-move, stock: (x.stock||0)+move } : x);
  save.items();
}


// ── ICONS + NAV ───────────────────────────────────────────────────────────────
const SvgIcon = ({d, d2, viewBox='0 0 24 24', stroke=true, size=18}) => (
  <svg width={size} height={size} viewBox={viewBox}
    fill={stroke?'none':'currentColor'} stroke={stroke?'currentColor':'none'}
    strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    <path d={d}/>{d2 && <path d={d2}/>}
  </svg>
);

const NAV = [
  {id:'dashboard',  label:'Dashboard',   icon:<SvgIcon d="M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h7v7h-7z"/>},
  {id:'sale',       label:'New Sale',    icon:<SvgIcon d="M12 5v14M5 12h14"/>},
  {id:'orders',     label:'Sales Order', icon:<SvgIcon d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2" d2="M9 5a2 2 0 012-2h2a2 2 0 012 2v0a2 2 0 01-2 2h-2a2 2 0 01-2-2v0z"/>},
  {id:'outstanding',label:'Outstanding', icon:<SvgIcon d="M12 2a10 10 0 110 20A10 10 0 0112 2z" d2="M12 6v6l4 2"/>},
  {id:'parties',    label:'Parties',     icon:<SvgIcon d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" d2="M9 11a4 4 0 100-8 4 4 0 000 8zm8 2c1.66 0 3 1.34 3 3v2"/>},
  {id:'items',      label:'Inventory',   icon:<SvgIcon d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16zM3.27 6.96L12 12.01l8.73-5.05M12 22.08V12"/>},
  {id:'accounts',   label:'Accounts',    icon:<SvgIcon d="M3 10h18M7 15h1m4 0h1m-7 4h12a3 3 0 003-3V8a3 3 0 00-3-3H6a3 3 0 00-3 3v8a3 3 0 003 3z"/>},
  {id:'reports',    label:'Reports',     icon:<SvgIcon d="M18 20V10M12 20V4M6 20v-6"/>},
  {id:'insights',   label:'Insights',    icon:<SvgIcon d="M3 3v18h18" d2="M18 9l-5 5-3-3-5 5"/>},
  {id:'sales-log',  label:'Sales Log',   icon:<SvgIcon d="M4 6h16M4 10h16M4 14h10"/>},
  {id:'deliveries', label:'Deliveries',  icon:<SvgIcon d="M1 3h15v13H1zM16 8h4l3 3v5h-7V8z" d2="M5.5 21a2.5 2.5 0 100-5 2.5 2.5 0 000 5zM18.5 21a2.5 2.5 0 100-5 2.5 2.5 0 000 5z"/>},
  {id:'purch-log',  label:'Purchases',   icon:<SvgIcon d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5" stroke={false}/>},
  {id:'settings',   label:'Settings',    icon:<SvgIcon d="M12 15a3 3 0 100-6 3 3 0 000 6z" d2="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/>},
];

// Curated 4-tab set for the mobile bottom tab bar — a deliberate subset of
// NAV (not NAV.slice(0,5)), in a fixed order chosen for what's reached most
// on a phone. Everything else in NAV stays reachable via the hamburger
// drawer; "New Sale" additionally has its own shortcut in the top bar.
const BOTTOM_NAV = ['dashboard','items','outstanding','accounts'].map(id=>NAV.find(n=>n.id===id));

// ── RESPONSIVE HELPERS ─────────────────────────────────────────────────────────
// Single source of truth for the mobile breakpoint — every component that
// needs to branch desktop/mobile calls this instead of rolling its own
// resize listener. Matches the breakpoint App() already used for the
// sidebar/bottom-tab-bar swap.
const MOBILE_BREAKPOINT = 768;
function useIsMobile(bp = MOBILE_BREAKPOINT){
  const [mobile,setMobile] = useState(()=>window.innerWidth < bp);
  useEffect(()=>{
    const onResize = () => setMobile(window.innerWidth < bp);
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  },[bp]);
  return mobile;
}

// ── SHARED UI ──────────────────────────────────────────────────────────────────
function Badge({v,color='blue'}){
  const m={blue:[C.blueSoft,C.blue],green:[C.greenSoft,C.green],red:[C.redSoft,C.red],
    amber:[C.amberSoft,C.amber],purple:[C.purpleSoft,C.purple],gray:['#F3F4F6',C.sub]};
  const[bg,tc]=m[color]||m.blue;
  return <span style={{background:bg,color:tc,fontSize:11,fontWeight:600,padding:'2px 9px',borderRadius:20,whiteSpace:'nowrap'}}>{v}</span>;
}

function Btn({children,onClick,variant='primary',size='md',icon,disabled,full}){
  const sz={sm:{p:'5px 12px',fs:12},md:{p:'8px 16px',fs:13},lg:{p:'10px 22px',fs:14}}[size]||{p:'8px 16px',fs:13};
  const v={primary:{bg:C.blue,tc:'#fff',bd:'none'},secondary:{bg:C.surf,tc:C.txt,bd:`1px solid ${C.bdr}`},
    ghost:{bg:'transparent',tc:C.sub,bd:'none'},success:{bg:C.greenSoft,tc:C.green,bd:`1px solid ${C.green}40`},
    danger:{bg:C.redSoft,tc:C.red,bd:`1px solid ${C.red}40`},amber:{bg:C.amberSoft,tc:C.amber,bd:`1px solid ${C.amber}40`}}[variant]||{bg:C.blue,tc:'#fff',bd:'none'};
  return <button onClick={onClick} disabled={disabled} style={{padding:sz.p,fontSize:sz.fs,background:v.bg,
    color:v.tc,border:v.bd,borderRadius:7,cursor:disabled?'not-allowed':'pointer',fontWeight:600,
    display:'inline-flex',alignItems:'center',gap:5,opacity:disabled?0.5:1,fontFamily:'inherit',
    width:full?'100%':'auto',justifyContent:'center'}}>
    {icon&&<span>{icon}</span>}{children}
  </button>;
}

// Fixed-size icon action button — used where several row actions (view/edit/
// return/delete) sit side by side, so columns stay aligned regardless of how
// many actions a given row shows (a text-label button's width varies with
// its word, an icon's doesn't).
function IconBtn({onClick,title,color=C.sub,bg='transparent',children}){
  return <button onClick={onClick} title={title} style={{width:26,height:26,flexShrink:0,display:'inline-flex',
    alignItems:'center',justifyContent:'center',borderRadius:6,border:`1px solid ${color}40`,background:bg,
    color,cursor:'pointer',fontSize:13,padding:0,fontFamily:'inherit',lineHeight:1}}>{children}</button>;
}

function Inp({id,label,value,onChange,placeholder,type='text',required,hint,autoFocus,autoComplete='off',onKeyDown}){
  return <div style={{display:'flex',flexDirection:'column',gap:3}}>
    {label&&<label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>
      {label}{required&&<span style={{color:C.red}}> *</span>}</label>}
    <input id={id} type={type} value={value||''} onChange={e=>onChange(e.target.value)}
      placeholder={placeholder} autoFocus={autoFocus} autoComplete={autoComplete} onKeyDown={onKeyDown}
      style={{padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,
        color:C.txt,outline:'none',fontFamily:'inherit',background:C.surf}}
      onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
    {hint&&<span style={{fontSize:11,color:C.muted}}>{hint}</span>}
  </div>;
}

function Sel({label,value,onChange,options,placeholder,required}){
  return <div style={{display:'flex',flexDirection:'column',gap:3}}>
    {label&&<label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>
      {label}{required&&<span style={{color:C.red}}> *</span>}</label>}
    <select value={value||''} onChange={e=>onChange(e.target.value)} style={{padding:'8px 12px',borderRadius:7,
      border:`1.5px solid ${C.bdr}`,fontSize:13,color:value?C.txt:C.muted,outline:'none',background:C.surf,fontFamily:'inherit'}}>
      {placeholder&&<option value="">{placeholder}</option>}
      {options.map(o=><option key={o.value||o} value={o.value||o}>{o.label||o}</option>)}
    </select>
  </div>;
}

function Modal({title,onClose,children,width=560}){
  const mobile = useIsMobile();
  return <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:1000,
    display:'flex',alignItems:'center',justifyContent:'center',padding:mobile?0:16}}
    onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
    <div style={{background:C.surf,borderRadius:mobile?0:12,width:'100%',maxWidth:mobile?'100%':width,
      height:mobile?'100%':'auto',maxHeight:mobile?'100%':'90vh',
      overflow:'auto',boxShadow:mobile?'none':'0 20px 60px rgba(0,0,0,0.3)'}}>
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'16px 22px',
        borderBottom:`1px solid ${C.bdr}`,position:'sticky',top:0,background:C.surf,zIndex:1}}>
        <h2 style={{margin:0,fontSize:16,fontWeight:700,color:C.txt}}>{title}</h2>
        <button onClick={onClose} style={{background:'none',border:'none',cursor:'pointer',
          fontSize:20,color:C.sub,padding:4,borderRadius:5}}>×</button>
      </div>
      <div style={{padding:22}}>{children}</div>
    </div>
  </div>;
}

function Grid2({children}){
  const mobile = useIsMobile();
  return <div style={{display:'grid',gridTemplateColumns:mobile?'1fr':'1fr 1fr',gap:14}}>{children}</div>;
}

function SectionHead({title,sub,action}){
  const mobile=useIsMobile();
  return <div style={{display:'flex',alignItems:mobile?'stretch':'flex-start',justifyContent:'space-between',flexWrap:'wrap',gap:mobile?10:0,marginBottom:mobile?14:18}}>
    <div style={{minWidth:0}}>
      <h2 style={{margin:0,fontSize:mobile?16:18,fontWeight:800,color:C.txt,letterSpacing:-.3}}>{title}</h2>
      {sub&&<p style={{margin:'3px 0 0',fontSize:12,color:C.sub}}>{sub}</p>}
    </div>
    {action&&<div style={mobile?{width:'100%'}:undefined}>{action}</div>}
  </div>;
}

function StatCard({label,value,sub,color=C.blue,icon,onClick}){
  return <div onClick={onClick} style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,
    borderTop:`3px solid ${color}`,padding:'16px 18px',cursor:onClick?'pointer':'default'}}>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start'}}>
      <div>
        <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.7,marginBottom:5}}>{label}</div>
        <div style={{fontSize:22,fontWeight:800,color,fontVariantNumeric:'tabular-nums',letterSpacing:-.5}}>{value}</div>
        {sub&&<div style={{fontSize:11,color:C.sub,marginTop:3}}>{sub}</div>}
      </div>
      {icon&&<span style={{fontSize:24,opacity:.2}}>{icon}</span>}
    </div>
  </div>;
}

function Chip({v,active,onClick,color}){
  return <button onClick={onClick} style={{padding:'5px 13px',borderRadius:20,fontFamily:'inherit',
    border:`1.5px solid ${active?(color||C.blue):'transparent'}`,
    background:active?((color||C.blue)+'12'):'#F3F4F6',
    color:active?(color||C.blue):C.sub,fontSize:12,fontWeight:600,cursor:'pointer',whiteSpace:'nowrap'}}>{v}</button>;
}

// Dual-rendering table: a verbatim desktop <table> (pass the exact markup
// you already have via the `columns`/`rows` contract below), or — below the
// mobile breakpoint — one bordered card per row with label:value lines.
// Call with `mobile` from useIsMobile() at the call site (not computed
// internally) so callers that already need the flag for other layout
// decisions don't end up with two separate resize listeners.
//
// columns: [{ key, label, cardLabel?, align, render(row,i)=>node, hideOnMobile? }]
// rows: array of row data objects
// keyFn: (row,i)=>string|number — React key, also passed to onRowClick/renderSubRow
// onRowClick?: (row)=>void — desktop <tr> onClick AND mobile card onClick
// renderSubRow?: (row)=>node|null — extra content rendered after the main
//    cells, inside the same <tr> colSpan block on desktop and inside the same
//    card on mobile (e.g. DamagedInventory's "Recovered from" breakdown)
// groupHeader?: (row, prevRow)=>node|null — returns a header element to
//    splice in before this row (desktop: full-width <tr><td colSpan>, mobile:
//    a plain heading above the card) when it differs from prevRow, e.g.
//    category separators in ItemMasters
// footer?: node — desktop: rendered as a <tfoot> row spanning all columns;
//    mobile: rendered as its own summary card after the last row
// emptyLabel?: string — shown when rows.length===0
function CardTable({columns,rows,keyFn,mobile,onRowClick,renderSubRow,groupHeader,emptyLabel='No records',footer,rowStyle}){
  const kf = keyFn || ((r,i)=>r.id ?? i);
  if(!rows || rows.length===0){
    return <div style={{padding:'30px 10px',textAlign:'center',color:C.sub,fontSize:13}}>{emptyLabel}</div>;
  }
  if(!mobile){
    return <table style={{width:'100%',borderCollapse:'collapse'}}>
      <thead>
        <tr style={{borderBottom:`1.5px solid ${C.bdr}`}}>
          {columns.map(c=><th key={c.key} style={{textAlign:c.align||'left',padding:'8px 10px',
            fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>{c.label}</th>)}
        </tr>
      </thead>
      <tbody>
        {rows.map((row,i)=>{
          const prev = i>0 ? rows[i-1] : null;
          const gh = groupHeader ? groupHeader(row,prev) : null;
          const sub = renderSubRow ? renderSubRow(row) : null;
          const rs = rowStyle ? rowStyle(row) : null;
          return <Fragment key={kf(row,i)}>
            {gh && <tr><td colSpan={columns.length} style={{padding:0}}>{gh}</td></tr>}
            <tr onClick={onRowClick?()=>onRowClick(row):undefined}
              style={{borderBottom:sub?'none':`1px solid ${C.bdr}`,cursor:onRowClick?'pointer':'default',...rs}}>
              {columns.map(c=><td key={c.key} style={{textAlign:c.align||'left',padding:'8px 10px',fontSize:13,color:C.txt}}>
                {c.render?c.render(row,i):row[c.key]}
              </td>)}
            </tr>
            {sub && <tr style={{borderBottom:`1px solid ${C.bdr}`}}><td colSpan={columns.length} style={{padding:0,...(rs?{background:rs.background}:{})}}>{sub}</td></tr>}
          </Fragment>;
        })}
      </tbody>
      {footer && <tfoot><tr><td colSpan={columns.length} style={{padding:0}}>{footer}</td></tr></tfoot>}
    </table>;
  }
  // Mobile: one card per row.
  return <div style={{display:'flex',flexDirection:'column',gap:10}}>
    {rows.map((row,i)=>{
      const prev = i>0 ? rows[i-1] : null;
      const gh = groupHeader ? groupHeader(row,prev) : null;
      const sub = renderSubRow ? renderSubRow(row) : null;
      const rs = rowStyle ? rowStyle(row) : null;
      const visibleCols = columns.filter(c=>!c.hideOnMobile);
      return <Fragment key={kf(row,i)}>
        {gh && <div>{gh}</div>}
        <div onClick={onRowClick?()=>onRowClick(row):undefined}
          style={{border:`1px solid ${C.bdr}`,borderRadius:9,padding:'12px 14px',background:C.surf,
            cursor:onRowClick?'pointer':'default',display:'flex',flexDirection:'column',gap:7,...rs}}>
          {visibleCols.map(c=><div key={c.key} style={{display:'flex',justifyContent:'space-between',
            alignItems:'flex-start',gap:10}}>
            <span style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',
              letterSpacing:.5,whiteSpace:'nowrap',flexShrink:0}}>{c.cardLabel||c.label}</span>
            <span style={{fontSize:13,color:C.txt,textAlign:'right'}}>{c.render?c.render(row,i):row[c.key]}</span>
          </div>)}
          {sub && <div style={{marginTop:2}}>{sub}</div>}
        </div>
      </Fragment>;
    })}
    {footer && <div style={{marginTop:2}}>{footer}</div>}
  </div>;
}

function useToast(){
  const[toasts,setToasts]=useState([]);
  const toast=useCallback((msg,type='success')=>{
    const id=Date.now();
    setToasts(t=>[...t,{id,msg,type}]);
    setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)),3000);
  },[]);
  const TC=()=><div style={{position:'fixed',top:20,right:20,zIndex:9999,display:'flex',flexDirection:'column',gap:8}}>
    {toasts.map(t=><div key={t.id} style={{padding:'11px 16px',borderRadius:9,fontSize:13,fontWeight:600,
      background:t.type==='error'?C.red:t.type==='warn'?C.amber:C.green,
      color:'#fff',boxShadow:'0 4px 20px rgba(0,0,0,0.2)',maxWidth:300}}>{t.msg}</div>)}
  </div>;
  return{toast,TC};
}

// First-letter shortcut for a category <select> — typing "p" jumps to Party
// Payment/Receipt, "v" to Vendor Payment, etc. Native <select> typeahead
// already does this by matching the option text, but our category options
// are rendered as "💰 Party Payment / Receipt" (icon + label), so the browser
// matches against the emoji, not the letter, and the shortcut silently does
// nothing — this matches against the label alone instead, always landing on
// the first category (in TXN_CATEGORIES order) whose label starts with that
// letter — e.g. "p" always lands on Party Payment/Receipt, never Petrol or
// Personal/Family, since it's declared first.
function useCategoryTypeahead(cats, onPick) {
  return e => {
    if (e.ctrlKey||e.metaKey||e.altKey||e.key.length!==1) return;
    const key = e.key.toLowerCase();
    if (!/[a-z0-9]/.test(key)) return;
    const match = cats.find(c=>(c.label||'').toLowerCase().startsWith(key));
    if (!match) return;
    e.preventDefault();
    onPick(match.id);
  };
}

// ── ICONS + NAV ───────────────────────────────────────────────────────────────
function OutstandingDetail({party,od,onClose}){
  const mobile=useIsMobile();
  const newSales=state.sales.filter(s=>s.partyId===party.id).reduce((s,t)=>s+t.total,0);
  const payments=state.payments.filter(p=>p.partyId===party.id).reduce((s,p)=>s+p.amount,0);
  const currentOD=od+newSales-payments;
  const[response,setResponse]=useState('');
  const[saved,setSaved]=useState(false);
  const[notRecoverable,setNotRecoverable]=useState(!!party.notRecoverable);
  const[phone,setPhone]=useState(party.phone||'');
  const[phEdit,setPhEdit]=useState(false);
  const[phVal,setPhVal]=useState(party.phone||'');
  const[discounts,setDiscounts]=useState({});   // {billKey: amountStr} — write-off amounts entered against each invoice
  const[repairs,setRepairs]=useState({});        // {billKey: amountStr} — repair-expenditure deduction per invoice
  const[repairExpOv,setRepairExpOv]=useState({}); // {billKey: amountStr} — expense-to-book override, defaults to repairs amount
  const{toast,TC}=useToast();
  const lastFU=useMemo(()=>[...state.followups].filter(f=>f.partyId===party.id).sort((a,b)=>b.date.localeCompare(a.date))[0],[]);
  // getPartyBills covers opening balance + historical + NEW in-app sales via
  // FIFO — the raw PARTY_INVOICES map only has historical rows, which left the
  // WhatsApp reminder's invoice breakdown blank for any party whose outstanding
  // comes from opening balance or new sales instead of a historical bill.
  const bills=useMemo(()=>getPartyBills(party).filter(b=>b.remaining>0.5),[party.id,state.sales.length,state.payments.length]);

  const saveResponse=()=>{
    if(!response.trim())return;
    const parsed=parseResponse(response);
    const fu={id:uid('FU'),partyId:party.id,partyName:party.name,date:today(),response,
      tags:parsed.tags,followUpDate:parsed.followUpDate,label:parsed.label,status:'open'};
    state.followups=[fu,...state.followups];save.followups();setSaved(true);
    toast('Saved'+(parsed.followUpDate?' · Reminder '+fmtD(parsed.followUpDate):''));
  };

  const invDetails=bills.map(b=>'• '+fmtD(b.date)+' — '+fmt(b.remaining)+' ('+b.daysSince+'d from invoice)').join('\n');
  const waUrl='https://wa.me/91'+phone+'?text='+encodeURIComponent(applyTemplate('payment',{partyName:party.name,amount:fmt(currentOD),invoiceDetails:invDetails}));
  const savePhone=()=>{
    if(!phVal||phVal.length<10)return toast('Enter a 10-digit number','error');
    updatePartyPhone(party.id,phVal);
    setPhone(phVal); setPhEdit(false); toast('Number updated ✓');
  };

  // Same write-off mechanics as the Accounts module's per-invoice discount
  // panel (a 'writeoff' txn + a linked payments record, no real cash account
  // involved) — surfaced here too so a discount can be recorded right where
  // the invoice is being reviewed, without a trip to Accounts.
  const discTotal=Object.values(discounts).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const doWriteOff=()=>{
    if(discTotal<=0)return;
    const woDate=today();
    const notes='Discount/write-off '+bills.filter(b=>parseFloat(discounts[b.key])>0).map(b=>fmtD(b.date)+' ₹'+discounts[b.key]).join(', ');
    const wo={id:uid('WO'),type:'writeoff',date:woDate,partyId:party.id,partyName:party.name,amount:discTotal,account:'writeoff',notes,createdAt:new Date().toISOString()};
    state.txns=[wo,...state.txns]; save.txns();
    const woPmt={id:uid('WOP'),txnId:wo.id,date:woDate,partyId:party.id,partyName:party.name,amount:discTotal,account:'writeoff',notes:'Discount write-off'};
    state.payments=[woPmt,...state.payments]; save.payments();
    setDiscounts({});
    toast('Wrote off '+fmt(discTotal)+' ✓');
  };

  // Repair expenditure: reduces the bill exactly like the write-off above
  // (same 'writeoff' account, no cash involved) but reimburses the party for
  // a repair they paid for themselves, and books a repairExpenses record per
  // bill so it's reported separately from plain discounts/write-offs.
  const repairTotal=Object.values(repairs).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const doRepairExpense=()=>{
    if(repairTotal<=0)return;
    const rDate=today();
    const notes='Repair expenditure '+bills.filter(b=>parseFloat(repairs[b.key])>0).map(b=>fmtD(b.date)+' ₹'+repairs[b.key]).join(', ');
    const wo={id:uid('WO'),type:'writeoff',date:rDate,partyId:party.id,partyName:party.name,amount:repairTotal,account:'writeoff',notes,createdAt:new Date().toISOString()};
    state.txns=[wo,...state.txns]; save.txns();
    const woPmt={id:uid('WOP'),txnId:wo.id,date:rDate,partyId:party.id,partyName:party.name,amount:repairTotal,account:'writeoff',notes:'Repair expenditure adjustment'};
    state.payments=[woPmt,...state.payments]; save.payments();
    bills.forEach(b=>{
      const amt=parseFloat(repairs[b.key])||0;
      if(amt<=0)return;
      const expAmt=parseFloat(repairExpOv[b.key])||amt;
      recordRepairExpense({partyId:party.id,partyName:party.name,date:rDate,itemId:'',itemName:'',qty:0,billedAmount:0,expenseAmount:expAmt,source:'payment',linkedTxnId:wo.id,notes:'Reimbursed against '+fmtD(b.date)+' invoice'});
    });
    setRepairs({});setRepairExpOv({});
    toast('Repair expenditure recorded — '+fmt(repairTotal)+' adjusted ✓');
  };

  return <Modal title={party.name} onClose={onClose} width={520}>
    <TC/>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
      <div>
        <div style={{fontSize:11,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:2}}>Outstanding</div>
        <div style={{fontSize:28,fontWeight:800,color:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(currentOD)}</div>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:8,alignItems:'flex-end'}}>
        {phEdit
          ?<div style={{display:'flex',gap:4,alignItems:'center'}}>
            <input value={phVal} onChange={e=>setPhVal(e.target.value)} placeholder="10-digit number" autoFocus autoComplete="off"
              onKeyDown={e=>{if(e.key==='Enter')savePhone();if(e.key==='Escape')setPhEdit(false);}}
              style={{width:130,padding:'6px 9px',borderRadius:6,border:`1.5px solid ${C.blue}`,fontSize:12,outline:'none',fontFamily:'inherit'}}/>
            <button onClick={savePhone} style={{padding:'6px 10px',borderRadius:6,background:C.blue,color:'#fff',border:'none',cursor:'pointer',fontSize:12,fontFamily:'inherit',fontWeight:600}}>✓</button>
            <button onClick={()=>setPhEdit(false)} style={{padding:'6px 8px',borderRadius:6,background:'none',color:C.muted,border:`1px solid ${C.bdr}`,cursor:'pointer',fontSize:12}}>✕</button>
          </div>
          :<div style={{display:'flex',gap:6,alignItems:'center'}}>
            {phone&&<a href={waUrl} target="_blank" rel="noreferrer"
              style={{padding:'8px 16px',background:'#25D366',color:'#fff',borderRadius:8,textDecoration:'none',fontWeight:700,fontSize:13}}>📲 WhatsApp</a>}
            <button onClick={()=>{setPhVal(phone);setPhEdit(true);}} title={phone?'Edit number':'Add number'}
              style={{padding:'6px 8px',borderRadius:6,border:`1px solid ${C.bdr}`,background:C.surf,color:C.sub,cursor:'pointer',fontSize:13}}>✎</button>
          </div>}
        {lastFU&&<span style={{fontSize:11,padding:'3px 8px',borderRadius:20,fontWeight:600,
          background:(TAG_COLORS[lastFU.tags[0]]||C.blue)+'18',color:TAG_COLORS[lastFU.tags[0]]||C.blue}}>
          Last: {lastFU.label} · {fmtD(lastFU.date)}
        </span>}
        <label style={{display:'flex',alignItems:'center',gap:6,fontSize:11,color:C.sub,cursor:'pointer'}}>
          <input type="checkbox" checked={notRecoverable} onChange={e=>{
            const v=e.target.checked;
            setNotRecoverable(v);
            state.parties=state.parties.map(x=>x.id===party.id?{...x,notRecoverable:v}:x);
            save.parties();
            toast(v?'Marked as Not Recoverable':'Unmarked ✓');
          }}/>
          Not Recoverable
        </label>
      </div>
    </div>

    {bills.length>0&&<div style={{marginBottom:16}}>
      <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>Outstanding Invoices</div>
      {bills.map(b=>(
        <div key={b.key} style={{padding:'10px 13px',borderRadius:8,marginBottom:6,
          border:`1.5px solid ${b.overdue>0?C.red+'50':C.bdr}`,background:b.overdue>0?C.redSoft:'#F6F8FA'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:3}}>
            <div style={{display:'flex',alignItems:'center',gap:8}}>
              <span style={{fontWeight:700,color:C.txt,fontSize:13}}>{fmtD(b.date)}</span>
              <span style={{fontSize:11,fontWeight:700,padding:'2px 8px',borderRadius:20,
                background:b.overdue>0?C.red:C.green,color:'#fff'}}>
                {b.overdue>0?b.overdue+'d overdue':'Within credit'}
              </span>
            </div>
            <span style={{fontWeight:800,color:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(b.remaining)}</span>
          </div>
          <div style={{fontSize:11,color:C.sub,marginBottom:6}}>Billed {fmt(b.total)} · {b.items||'—'}</div>
          <div style={{display:'flex',alignItems:mobile?'flex-start':'center',gap:mobile?8:6,flexWrap:'wrap',flexDirection:mobile?'column':'row'}}>
            <div style={{display:'flex',alignItems:'center',gap:6}}>
              <span style={{fontSize:11,color:C.sub}}>Write off ₹</span>
              <input type="number" placeholder="0" value={discounts[b.key]||''}
                onChange={e=>setDiscounts(d=>({...d,[b.key]:e.target.value}))}
                style={{width:70,padding:'4px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,outline:'none',fontFamily:'inherit',textAlign:'right'}}
                onFocus={e=>e.target.style.borderColor=C.amber} onBlur={e=>e.target.style.borderColor=C.bdr}/>
            </div>
            <div style={{display:'flex',alignItems:'center',gap:6,flexWrap:'wrap'}}>
              <span style={{fontSize:11,color:C.sub,marginLeft:mobile?0:10}}>Repair ₹</span>
              <input type="number" placeholder="0" value={repairs[b.key]||''}
                onChange={e=>setRepairs(d=>({...d,[b.key]:e.target.value}))}
                style={{width:70,padding:'4px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,outline:'none',fontFamily:'inherit',textAlign:'right'}}
                onFocus={e=>e.target.style.borderColor=C.purple} onBlur={e=>e.target.style.borderColor=C.bdr}/>
              {parseFloat(repairs[b.key])>0&&<>
                <span style={{fontSize:10,color:C.muted}}>↳ Expense ₹</span>
                <input type="number" placeholder={repairs[b.key]} value={repairExpOv[b.key]||''}
                  onChange={e=>setRepairExpOv(d=>({...d,[b.key]:e.target.value}))}
                  style={{width:60,padding:'3px 6px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:11,outline:'none',fontFamily:'inherit',textAlign:'right',color:C.muted}}
                  onFocus={e=>e.target.style.borderColor=C.purple} onBlur={e=>e.target.style.borderColor=C.bdr}/>
              </>}
            </div>
            {(parseFloat(discounts[b.key])>0||parseFloat(repairs[b.key])>0)&&<span style={{fontSize:11,color:C.amber,width:'100%'}}>→ net {fmt(Math.max(0,b.remaining-(parseFloat(discounts[b.key])||0)-(parseFloat(repairs[b.key])||0)))}</span>}
          </div>
        </div>
      ))}
      {discTotal>0&&<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',
        padding:'8px 4px 0'}}>
        <span style={{fontSize:12,color:C.amber,fontWeight:600}}>Total write-off: {fmt(discTotal)}</span>
        <Btn onClick={doWriteOff} size="sm">Write Off {fmt(discTotal)}</Btn>
      </div>}
      {repairTotal>0&&<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',
        padding:'8px 4px 0'}}>
        <span style={{fontSize:12,color:C.purple,fontWeight:600}}>Total repair expenditure: {fmt(repairTotal)}</span>
        <Btn onClick={doRepairExpense} size="sm">Record Repair Expenditure {fmt(repairTotal)}</Btn>
      </div>}
    </div>}

    <div style={{borderTop:`1px solid ${C.bdr}`,paddingTop:16}}>
      <div style={{fontSize:12,fontWeight:700,color:C.txt,marginBottom:8}}>Capture Response</div>
      {!saved
        ?<>
          <textarea value={response} onChange={e=>setResponse(e.target.value)} rows={3}
            placeholder="What did they say? e.g. Will pay Monday / Already paid / Busy 2 days"
            style={{width:'100%',padding:'10px 12px',borderRadius:8,border:`1.5px solid ${C.bdr}`,
              fontSize:13,fontFamily:'inherit',resize:'none',outline:'none',boxSizing:'border-box',color:C.txt}}
            onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
          {response.trim()&&(()=>{const p=parseResponse(response);return <div style={{marginTop:6,display:'flex',gap:6,flexWrap:'wrap',alignItems:'center'}}>
            <span style={{fontSize:11,color:C.sub}}>Tag:</span>
            {p.tags.map(t=><span key={t} style={{fontSize:11,fontWeight:700,padding:'2px 9px',borderRadius:20,background:(TAG_COLORS[t]||C.blue)+'18',color:TAG_COLORS[t]||C.blue}}>{t}</span>)}
            {p.followUpDate&&<span style={{fontSize:11,color:C.green,fontWeight:600}}>📅 {fmtD(p.followUpDate)}</span>}
          </div>;})()} 
          <div style={{display:'flex',justifyContent:'flex-end',marginTop:10}}>
            <Btn onClick={saveResponse} size="sm">Save Response</Btn>
          </div>
        </>
        :<div style={{padding:'10px 14px',background:C.greenSoft,borderRadius:8,fontSize:13,color:C.green,fontWeight:600}}>✓ Response saved</div>
      }
    </div>
  </Modal>;
}

// ── OUTSTANDING ───────────────────────────────────────────────────────────────
function Outstanding(){
  const mobile=useIsMobile();
  const[routeF,setRouteF]=useState('');   // '' | 'CG North' | 'CG South' | 'Odisha' | 'Raipur' | 'Lighter'
  const[ageF,setAgeF]    =useState('');   // '' | '0-30' | '30-60' | '60+'
  const[profile,setProfile]=useState(null);
  const[search,setSearch]=useState('');
  const[detail,setDetail]=useState(null);
  const[cityView,setCityView]=useState(false);
  const[showCityOrder,setShowCityOrder]=useState(false);

  // All parties with OD (full unfiltered list for summary cards)
  const allOD=useMemo(()=>{
    return state.parties.filter(p=>p.type==='customer').map(p=>{
      // getPartyBills does FIFO payment allocation against the oldest bill first
      // and computes overdue from the real current date — the old static
      // PARTY_INVOICES[3] day-count was frozen at whenever the historical data
      // was generated and never accounted for payments made since, so it drifted
      // stale (wrong bucket) as real time passed and payments came in.
      const bills=getPartyBills(p).filter(b=>b.remaining>0.5);
      const od=bills.reduce((s,b)=>s+b.remaining,0);
      // getPartyBills stamps any opening-balance excess with a synthetic
      // 1900-01-01 date so it sorts first for FIFO payment purposes — that's
      // fine for payment order, but "days since 1900" is nonsense for the age
      // bucket, so pick the oldest bill with a real invoice date instead.
      const realBills=bills.filter(b=>b.date!=='1900-01-01');
      // Split outstanding by EACH bill's own overdue status, not the party's
      // oldest bill — a party with one within-credit invoice and one overdue
      // invoice must contribute to both buckets separately, otherwise the
      // whole balance gets dumped into whichever bucket the oldest bill lands
      // in and the age-bucket totals overstate one bucket at the other's expense.
      const withinAmt=bills.filter(b=>b.overdue<=0).reduce((s,b)=>s+b.remaining,0);
      const overdue030Amt=bills.filter(b=>b.overdue>0&&b.overdue<=30).reduce((s,b)=>s+b.remaining,0);
      const overdue31Amt=bills.filter(b=>b.overdue>30).reduce((s,b)=>s+b.remaining,0);
      return{...p,od,overdueDays:realBills.length?realBills[0].overdue:0,withinAmt,overdue030Amt,overdue31Amt,isLighter:p.credit===3,notRecoverable:!!p.notRecoverable};
    }).filter(p=>p.od>0);
  },[state.payments.length,state.sales.length,state.txns.length,state.parties]);

  // Which per-bill bucket amount a given ageF corresponds to, if any (the
  // three amount buckets only — Not Recoverable isn't bill-split).
  const isAmountBucket = f => f==='within-credit'||f==='0-30'||f==='31+';
  const bucketAmt = (p,f) => f==='within-credit'?p.withinAmt:f==='0-30'?p.overdue030Amt:p.overdue31Amt;

  // Filtered list (applies route + age + search filters). A party matches an
  // amount bucket if ANY of its bills fall in that bucket — the same party
  // can show up under more than one bucket when its invoices span both.
  const odList=useMemo(()=>allOD.filter(p=>{
    const routeMatch = !routeF ||(routeF==='Lighter'?p.isLighter&&!p.notRecoverable:p.route===routeF&&!p.isLighter&&!p.notRecoverable);
    const ageMatch   = !ageF  ||(ageF==='not-recoverable'?p.notRecoverable:!p.notRecoverable&&bucketAmt(p,ageF)>0.5);
    const srchMatch  = !search||p.name.toLowerCase().includes(search.toLowerCase())||p.city?.toLowerCase().includes(search.toLowerCase());
    return routeMatch&&ageMatch&&srchMatch;
  }).sort((a,b)=>b.overdueDays-a.overdueDays||b.od-a.od),[allOD,routeF,ageF,search]);

  // When an amount bucket is active, sum that bucket's portion (not the
  // party's full OD) so this total matches the bucket card above it.
  const total=useMemo(()=>{
    if(isAmountBucket(ageF)) return odList.reduce((s,p)=>s+bucketAmt(p,ageF),0);
    return odList.reduce((s,p)=>s+p.od,0);
  },[odList,ageF]);

  // Summary by segment (always from allOD, unfiltered) — route/lighter only;
  // Not Recoverable lives in the age-bucket row below instead (2nd filter level).
  const segments=useMemo(()=>{
    const regular=allOD.filter(p=>!p.isLighter&&!p.notRecoverable);
    const lighter=allOD.filter(p=>p.isLighter&&!p.notRecoverable);
    const byRoute=ROUTES.map(r=>{
      const ps=regular.filter(p=>p.route===r);
      return{id:r,label:r,total:ps.reduce((s,p)=>s+p.od,0),count:ps.length,
        color:r==='CG North'?C.blue:r==='CG South'?C.green:r==='Odisha'?C.amber:C.purple};
    }).filter(x=>x.count>0);
    const lighterSeg={id:'Lighter',label:'Lighter Dues',total:lighter.reduce((s,p)=>s+p.od,0),count:lighter.length,color:'#7C3AED'};
    const segs=[...byRoute];
    if(lighterSeg.count>0)segs.push(lighterSeg);
    return segs;
  },[allOD]);

  const grandTotal=allOD.reduce((s,p)=>s+p.od,0);

  // Scope the age-bucket row to whichever route/Lighter card is selected up
  // top, so the breakup reflects that region instead of always the global
  // total. Not Recoverable still needs a region's own not-recoverable parties
  // included here (they're excluded from the route cards above), so this
  // matches purely on region/Lighter, not on the notRecoverable flag.
  const regionOD=useMemo(()=>{
    if(!routeF) return allOD;
    if(routeF==='Lighter') return allOD.filter(p=>p.isLighter);
    return allOD.filter(p=>p.route===routeF&&!p.isLighter);
  },[allOD,routeF]);

  // Second-level filter row: within-credit / 0-30 overdue / 31+ overdue / Not
  // Recoverable — mutually exclusive with each other and with the route row.
  // The three amount buckets sum each party's bill-level split (withinAmt /
  // overdue030Amt / overdue31Amt) rather than the whole party OD, so a party
  // with one within-credit invoice and one overdue invoice contributes the
  // right amount to each bucket instead of dumping its whole balance into one.
  const ageBuckets=useMemo(()=>{
    const recoverable=regionOD.filter(p=>!p.notRecoverable);
    const notRecoverablePs=regionOD.filter(p=>p.notRecoverable);
    return [
      {id:'within-credit',  label:'Within Credit Period', color:C.green,
        total:recoverable.reduce((s,p)=>s+p.withinAmt,0),    count:recoverable.filter(p=>p.withinAmt>0.5).length},
      {id:'0-30',           label:'0–30 Days Overdue',    color:C.amber,
        total:recoverable.reduce((s,p)=>s+p.overdue030Amt,0),count:recoverable.filter(p=>p.overdue030Amt>0.5).length},
      {id:'31+',            label:'31+ Days Overdue',     color:C.red,
        total:recoverable.reduce((s,p)=>s+p.overdue31Amt,0), count:recoverable.filter(p=>p.overdue31Amt>0.5).length},
      {id:'not-recoverable',label:'🚫 Not Recoverable',   color:C.muted,
        total:notRecoverablePs.reduce((s,p)=>s+p.od,0),      count:notRecoverablePs.length},
    ];
  },[regionOD]);

  const setRouteFilter = r => { setRouteF(routeF===r?'':r); setAgeF(''); };
  const setAgeFilter   = a => { setAgeF(ageF===a?'':a);     };

  // City breakdown — grouped from whatever's currently filtered (odList), so
  // it respects the route/age filters already on screen rather than always
  // showing every city. Sorted highest-total first: when visiting a city for
  // collection, that's the order worth checking off first.
  const partyAmt = p => isAmountBucket(ageF)?bucketAmt(p,ageF):p.od;
  // Cities the user has manually ordered (state.cityOrder, set via
  // CityOrderModal) come first in that exact sequence — matching the actual
  // route walked on a collection trip — with any city not yet placed in the
  // route falling after, ranked by total due like before.
  const cityRouteIdx = city => { const i=(state.cityOrder||[]).indexOf(city); return i===-1?Infinity:i; };
  const byCity=useMemo(()=>{
    const map={};
    for(const p of odList){
      const city=p.city||'Unknown';
      if(!map[city])map[city]={city,parties:[],total:0};
      map[city].parties.push(p);
      map[city].total+=partyAmt(p);
    }
    return Object.values(map).map(c=>({...c,parties:[...c.parties].sort((a,b)=>partyAmt(b)-partyAmt(a))}))
      .sort((a,b)=>{
        const ia=cityRouteIdx(a.city),ib=cityRouteIdx(b.city);
        return ia!==ib?ia-ib:b.total-a.total;
      });
  },[odList,ageF,state.cityOrder]);

  // Plain-text collection lists, meant to be shared with whoever's doing the
  // rounds (self, a collection agent) rather than sent to one specific
  // customer — wa.me with no phone number opens WhatsApp's own contact/group
  // picker instead of presetting a recipient.
  const cityMsg = c => {
    const lines=c.parties.map((p,i)=>(i+1)+'. '+p.name+(p.city?'':'')+' — '+fmt(partyAmt(p))+(p.overdueDays>0?' ('+p.overdueDays+'d overdue)':'')).join('\n');
    return '📍 *'+c.city+'* — Collection List\nTotal: *'+fmt(c.total)+'* from '+c.parties.length+(c.parties.length===1?' party':' parties')+'\n\n'+lines;
  };
  const shareCity = c => window.open('https://wa.me/?text='+encodeURIComponent(cityMsg(c)),'_blank');
  const shareAllCities = () => {
    const grand=byCity.reduce((s,c)=>s+c.total,0);
    const header='🧾 *Outstanding Collection List*\nTotal: *'+fmt(grand)+'* across '+byCity.length+(byCity.length===1?' city':' cities');
    const body=byCity.map(cityMsg).join('\n\n──────────\n\n');
    window.open('https://wa.me/?text='+encodeURIComponent(header+'\n\n'+body),'_blank');
  };

  return <div>
    {profile&&<PartyLedger party={profile} onClose={()=>setProfile(null)}/>}
    {detail&&<OutstandingDetail party={detail.party} od={detail.od} onClose={()=>setDetail(null)}/>}
    {showCityOrder&&<CityOrderModal cities={byCity.map(c=>c.city)} onClose={()=>setShowCityOrder(false)}/>}

    <SectionHead title="Outstanding"
      sub={(routeF||ageF)?`${odList.length} parties · ${fmtL(total)} (filtered)`:  `${allOD.length} parties · ${fmtL(grandTotal)}`}/>

    {/* Route + Lighter summary cards — clickable */}
    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:10,marginBottom:14}}>
      {/* Grand total card */}
      <button onClick={()=>{setRouteF('');setAgeF('');setSearch('');}} style={{
        padding:'13px 16px',borderRadius:8,cursor:'pointer',fontFamily:'inherit',textAlign:'left',
        border:`2px solid ${!routeF&&!ageF?C.red:C.bdr}`,
        background:!routeF&&!ageF?C.redSoft:C.surf,transition:'all .15s'}}>
        <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:4}}>Total Due</div>
        <div style={{fontSize:18,fontWeight:800,color:C.red,fontVariantNumeric:'tabular-nums'}}>{fmtL(grandTotal)}</div>
        <div style={{fontSize:11,color:C.muted,marginTop:2}}>{allOD.length} parties</div>
      </button>
      {/* Per-segment cards */}
      {segments.map(seg=>(
        <button key={seg.id} onClick={()=>setRouteFilter(seg.id)} style={{
          padding:'13px 16px',borderRadius:8,cursor:'pointer',fontFamily:'inherit',textAlign:'left',
          border:`2px solid ${routeF===seg.id?seg.color:C.bdr}`,
          background:routeF===seg.id?seg.color+'12':C.surf,transition:'all .15s'}}>
          <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:4}}>
            {seg.id==='Lighter'?'🕯 Lighter':seg.label}
          </div>
          <div style={{fontSize:18,fontWeight:800,color:seg.color,fontVariantNumeric:'tabular-nums'}}>{fmtL(seg.total)}</div>
          <div style={{fontSize:11,color:C.muted,marginTop:2}}>{seg.count} parties</div>
        </button>
      ))}
    </div>

    {/* Aging bucket cards — clickable (2nd-level filter: within-credit / 0-30 / 31+ / Not Recoverable) */}
    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:10,marginBottom:14}}>
      {ageBuckets.map(b=>(
        <button key={b.id} onClick={()=>setAgeFilter(b.id)} style={{
          padding:'12px 16px',borderRadius:8,cursor:'pointer',fontFamily:'inherit',textAlign:'left',
          border:`2px solid ${ageF===b.id?b.color:C.bdr}`,
          background:ageF===b.id?b.color+'12':C.surf,transition:'all .15s'}}>
          <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:4}}>{b.label}</div>
          <div style={{fontSize:18,fontWeight:800,color:b.color,fontVariantNumeric:'tabular-nums'}}>{fmtL(b.total)}</div>
          <div style={{fontSize:11,color:C.muted,marginTop:2}}>{b.count} parties</div>
        </button>
      ))}
    </div>

    {/* Active filter banner */}
    {(routeF||ageF)&&<div style={{display:'flex',alignItems:mobile?'flex-start':'center',gap:8,marginBottom:10,padding:'8px 14px',
      background:C.blueSoft,borderRadius:8,border:`1px solid ${C.blue}20`,fontSize:12,flexDirection:mobile?'column':'row'}}>
      <div style={{display:'flex',alignItems:'center',gap:8,flexWrap:'wrap'}}>
        <span style={{color:C.blue,fontWeight:600}}>Showing:</span>
        {routeF&&<span style={{background:C.blue,color:'#fff',padding:'2px 9px',borderRadius:20,fontSize:11,fontWeight:600}}>{routeF}</span>}
        {ageF&&<span style={{background:C.amber,color:'#fff',padding:'2px 9px',borderRadius:20,fontSize:11,fontWeight:600}}>{ageF}</span>}
      </div>
      <button onClick={()=>{setRouteF('');setAgeF('');}} style={{marginLeft:mobile?0:'auto',alignSelf:mobile?'flex-end':'auto',background:'none',border:'none',cursor:'pointer',fontSize:12,color:C.sub,fontFamily:'inherit'}}>Clear filters ✕</button>
    </div>}

    {/* Search */}
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'8px 12px',marginBottom:12}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search party or city…"
        style={{width:'100%',border:'none',outline:'none',fontSize:13,color:C.txt,fontFamily:'inherit',background:'transparent'}}/>
    </div>

    {/* By Party / By City toggle + consolidated share — city view exists for
        collection trips: check every party in a city off in one place, and
        share the whole list (or just one city's) to WhatsApp beforehand. */}
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12,flexWrap:'wrap',gap:8}}>
      <div style={{display:'flex',gap:8}}>
        <Chip v="By Party" active={!cityView} onClick={()=>setCityView(false)}/>
        <Chip v="By City" active={cityView} onClick={()=>setCityView(true)}/>
      </div>
      {cityView&&byCity.length>0&&<div style={{display:'flex',gap:8}}>
        <button onClick={()=>setShowCityOrder(true)} style={{display:'inline-flex',alignItems:'center',gap:6,padding:'7px 14px',
          borderRadius:7,border:`1.5px solid ${C.bdr}`,color:C.sub,background:C.surf,cursor:'pointer',fontSize:12,fontWeight:700,fontFamily:'inherit'}}>
          ↕️ Reorder Route
        </button>
        <button onClick={shareAllCities} style={{display:'inline-flex',alignItems:'center',gap:6,padding:'7px 14px',
          borderRadius:7,border:'1.5px solid #25D366',color:'#15803D',background:'#F0FDF4',cursor:'pointer',fontSize:12,fontWeight:700,fontFamily:'inherit'}}>
          📲 Share All Cities (Consolidated)
        </button>
      </div>}
    </div>

    {cityView?
    /* City-grouped view */
    <div style={{display:'flex',flexDirection:'column',gap:14}}>
      {byCity.length===0&&<div style={{padding:'40px',textAlign:'center',color:C.muted,background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`}}>No parties match the filter</div>}
      {byCity.map(c=>(
        <div key={c.city} style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'10px 14px',background:'#F6F8FA',borderBottom:`1px solid ${C.bdr}`}}>
            <div>
              <span style={{fontWeight:700,color:C.txt,fontSize:14}}>📍 {c.city}</span>
              <span style={{fontSize:12,color:C.sub,marginLeft:8}}>{c.parties.length} {c.parties.length===1?'party':'parties'} · <b style={{color:C.red}}>{fmtL(c.total)}</b></span>
            </div>
            <button onClick={()=>shareCity(c)} style={{display:'inline-flex',alignItems:'center',gap:5,padding:'5px 12px',
              borderRadius:6,border:'1.5px solid #25D366',color:'#15803D',background:'#F0FDF4',cursor:'pointer',fontSize:11,fontWeight:700,fontFamily:'inherit'}}>
              📲 Share
            </button>
          </div>
          <CardTable mobile={mobile} rows={c.parties} keyFn={p=>p.id} onRowClick={p=>setDetail({party:p,od:p.ob})}
            rowStyle={p=>({background:p.isLighter?'#FAF5FF':p.overdueDays>90?C.redSoft:p.overdueDays>30?C.amberSoft:'transparent'})}
            columns={[
              {key:'idx',label:'#',render:(p,i)=><span style={{color:C.muted,fontSize:11}}>{i+1}</span>,hideOnMobile:true},
              {key:'party',label:'Party',render:p=>{
                const lastFU=state.followups.find(f=>f.partyId===p.id&&f.status==='open');
                return <div onClick={e=>{e.stopPropagation();setProfile(p);}}>
                  <div style={{fontWeight:600,color:C.blue,cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted',fontSize:13}}>{p.name}</div>
                  <div style={{display:'flex',gap:6,marginTop:2,flexWrap:'wrap'}}>
                    {p.isLighter&&<span style={{fontSize:10,background:'#FAF5FF',color:'#7C3AED',fontWeight:700,padding:'1px 7px',borderRadius:20,border:'1px solid #DDD6FE'}}>Lighter</span>}
                    {p.notRecoverable&&<span style={{fontSize:10,background:'#F3F4F6',color:C.muted,fontWeight:700,padding:'1px 7px',borderRadius:20,border:`1px solid ${C.bdr}`}}>Not Recoverable</span>}
                    {lastFU&&<span style={{fontSize:10,color:TAG_COLORS[lastFU.tags[0]]||C.blue,fontWeight:600}}>{lastFU.label}</span>}
                  </div>
                </div>;
              }},
              {key:'overdue',label:'Overdue',render:p=>p.overdueDays>0
                ?<span style={{fontWeight:700,fontSize:12,color:p.overdueDays>90?C.red:p.overdueDays>30?C.amber:C.sub}}>{p.overdueDays}d overdue</span>
                :<span style={{fontSize:12,color:p.isLighter?C.purple:C.green,fontWeight:600}}>{p.isLighter?'3d credit':'In credit'}</span>},
              {key:'amt',label:'Outstanding',align:'right',render:p=>
                <span style={{fontWeight:800,color:p.isLighter?'#7C3AED':C.red,fontVariantNumeric:'tabular-nums',fontSize:14,whiteSpace:'nowrap'}}>{fmt(partyAmt(p))}</span>},
            ]}/>
        </div>
      ))}
    </div>
    :<>
    {/* Party table */}
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <CardTable mobile={mobile} rows={odList} keyFn={(p,i)=>p.id||i} emptyLabel="No parties match the filter"
        onRowClick={p=>setDetail({party:p,od:p.ob})}
        rowStyle={p=>({background:p.isLighter?'#FAF5FF':p.overdueDays>90?C.redSoft:p.overdueDays>30?C.amberSoft:'transparent'})}
        columns={[
          {key:'idx',label:'#',render:(p,i)=><span style={{color:C.muted,fontSize:11}}>{i+1}</span>,hideOnMobile:true},
          {key:'party',label:'Party',render:p=>{
            const lastFU=state.followups.find(f=>f.partyId===p.id&&f.status==='open');
            return <div onClick={e=>{e.stopPropagation();setProfile(p);}}>
              <div style={{fontWeight:600,color:C.blue,cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted'}}>{p.name}</div>
              <div style={{display:'flex',gap:6,marginTop:2,flexWrap:'wrap'}}>
                {p.isLighter&&<span style={{fontSize:10,background:'#FAF5FF',color:'#7C3AED',fontWeight:700,padding:'1px 7px',borderRadius:20,border:'1px solid #DDD6FE'}}>Lighter</span>}
                {p.notRecoverable&&<span style={{fontSize:10,background:'#F3F4F6',color:C.muted,fontWeight:700,padding:'1px 7px',borderRadius:20,border:`1px solid ${C.bdr}`}}>Not Recoverable</span>}
                {lastFU&&<span style={{fontSize:10,color:TAG_COLORS[lastFU.tags[0]]||C.blue,fontWeight:600}}>{lastFU.label}</span>}
              </div>
            </div>;
          }},
          {key:'city',label:'City',render:p=><span style={{color:C.sub,fontSize:12}}>{p.city||'—'}</span>},
          {key:'route',label:'Route',render:p=><Badge v={p.isLighter?'Raipur':p.route||'—'} color={p.route==='CG North'?'blue':p.route==='CG South'?'green':p.route==='Odisha'?'amber':'gray'}/>},
          {key:'overdue',label:'Overdue',render:p=>p.overdueDays>0
            ?<span style={{fontWeight:700,fontSize:12,color:p.overdueDays>90?C.red:p.overdueDays>30?C.amber:C.sub}}>{p.overdueDays}d overdue</span>
            :<span style={{fontSize:12,color:p.isLighter?C.purple:C.green,fontWeight:600}}>{p.isLighter?'3d credit':'In credit'}</span>},
          {key:'amt',label:'Outstanding',align:'right',render:p=><>
            <span style={{fontWeight:800,color:p.isLighter?'#7C3AED':C.red,fontVariantNumeric:'tabular-nums',fontSize:14}}>{fmt(isAmountBucket(ageF)?bucketAmt(p,ageF):p.od)}</span>
            {isAmountBucket(ageF)&&p.od!==bucketAmt(p,ageF)&&<div style={{fontSize:10,fontWeight:600,color:C.muted}}>of {fmt(p.od)}</div>}
          </>},
          {key:'wa',label:'WA',align:'center',render:p=>{
            const pBills=getPartyBills(p).filter(b=>b.remaining>0.5);
            const pInvDetails=pBills.map(b=>'• '+fmtD(b.date)+' — '+fmt(b.remaining)+' ('+b.daysSince+'d from invoice)').join('\n');
            const waMsg=applyTemplate('payment',{partyName:p.name,amount:fmt(p.od),invoiceDetails:pInvDetails});
            return <div onClick={e=>e.stopPropagation()}>
              {p.phone
                ?<a href={'https://wa.me/91'+p.phone+'?text='+encodeURIComponent(waMsg)} target="_blank" rel="noreferrer" title={'Send to '+p.phone}
                    style={{display:'inline-flex',alignItems:'center',gap:5,padding:'5px 10px',
                      borderRadius:6,border:'1.5px solid #25D366',color:'#15803D',
                      textDecoration:'none',fontSize:12,fontWeight:700,background:'#F0FDF4'}}>
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="currentColor">
                      <path d="M17.47 14.38c-.3-.15-1.77-.87-2.04-.97-.27-.1-.47-.15-.67.15-.2.3-.77.97-.94 1.17-.17.2-.35.22-.65.07-.3-.15-1.26-.46-2.4-1.47-.89-.79-1.49-1.76-1.66-2.06-.17-.3-.02-.46.13-.61.13-.13.3-.35.45-.52.15-.17.2-.3.3-.5.1-.2.05-.37-.02-.52-.07-.15-.67-1.62-.92-2.22-.24-.58-.49-.5-.67-.51-.17-.01-.37-.01-.57-.01-.2 0-.52.07-.79.37-.27.3-1.04 1.01-1.04 2.47 0 1.46 1.07 2.87 1.22 3.07.15.2 2.1 3.2 5.08 4.49.71.31 1.26.5 1.69.64.71.22 1.36.19 1.87.11.57-.08 1.77-.72 2.02-1.42.25-.7.25-1.3.17-1.42-.07-.13-.27-.2-.57-.35z"/>
                      <path d="M12 2C6.48 2 2 6.48 2 12c0 1.85.5 3.58 1.37 5.07L2 22l5.07-1.35C8.5 21.5 10.22 22 12 22c5.52 0 10-4.48 10-10S17.52 2 12 2z" fillRule="evenodd"/>
                    </svg>WA
                  </a>
                :<span style={{fontSize:11,color:C.muted}}>—</span>}
            </div>;
          }},
        ]}/>
    </div>
    </>}
  </div>;
}

// Lets the user set the order cities appear in throughout the city-wise
// Outstanding view (and its WhatsApp shares) to match the actual route they
// walk on a collection trip, rather than always ranking by amount due.
// Persisted in state.cityOrder (a plain array of city names) — cities not
// yet placed keep falling back to amount-desc order, so a newly-appearing
// city is never lost, just unordered until added here.
function CityOrderModal({cities,onClose}){
  const[order,setOrder]=useState(()=>{
    const saved=(state.cityOrder||[]).filter(c=>cities.includes(c));
    const rest=cities.filter(c=>!saved.includes(c));
    return [...saved,...rest];
  });
  const move=(i,dir)=>{
    const j=i+dir;
    if(j<0||j>=order.length)return;
    const next=[...order];
    [next[i],next[j]]=[next[j],next[i]];
    setOrder(next);
  };
  const doSave=()=>{
    state.cityOrder=order;
    save.cityOrder();
    onClose?.();
  };
  return <Modal title="Reorder Cities — Collection Route" onClose={onClose} width={440}>
    <div style={{fontSize:12,color:C.sub,marginBottom:12}}>Set the order you visit cities for collection. Applies to the city-wise Outstanding list and its WhatsApp shares.</div>
    <div style={{maxHeight:380,overflowY:'auto',border:`1px solid ${C.bdr}`,borderRadius:8}}>
      {order.map((c,i)=><div key={c} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'9px 12px',borderBottom:i<order.length-1?`1px solid ${C.bdr}`:'none'}}>
        <div style={{display:'flex',alignItems:'center',gap:8}}>
          <span style={{fontSize:11,color:C.muted,width:18,textAlign:'right'}}>{i+1}.</span>
          <span style={{fontWeight:600,fontSize:13,color:C.txt}}>{c}</span>
        </div>
        <div style={{display:'flex',gap:4}}>
          <button onClick={()=>move(i,-1)} disabled={i===0} style={{width:26,height:26,borderRadius:6,border:`1px solid ${C.bdr}`,background:i===0?'#F6F8FA':C.surf,color:i===0?C.muted:C.txt,cursor:i===0?'default':'pointer',fontFamily:'inherit'}}>↑</button>
          <button onClick={()=>move(i,1)} disabled={i===order.length-1} style={{width:26,height:26,borderRadius:6,border:`1px solid ${C.bdr}`,background:i===order.length-1?'#F6F8FA':C.surf,color:i===order.length-1?C.muted:C.txt,cursor:i===order.length-1?'default':'pointer',fontFamily:'inherit'}}>↓</button>
        </div>
      </div>)}
    </div>
    <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:14}}>
      <Btn onClick={onClose} variant="secondary">Cancel</Btn>
      <Btn onClick={doSave}>Save Order</Btn>
    </div>
  </Modal>;
}

// ── PARTY INFO CARD ───────────────────────────────────────────────────────────
function PartyInfoCard({party}){
  const mobile=useIsMobile();
  const st=gPS(party.name);
  const ns=state.sales.filter(s=>s.partyId===party.id).reduce((s,t)=>s+t.total,0);
  const od=getCustomerBal(party.id);
  const tp=(st?.ts||0)+ns;
  const grade=od>100000?'D':od>50000?'C':od<=0?'A':'B';
  const gc={A:C.green,B:C.blue,C:C.amber,D:C.red};
  return <div style={{display:'grid',gridTemplateColumns:mobile?'repeat(2,1fr)':'repeat(4,1fr)',gap:8,marginTop:10,marginBottom:2}}>
    {[['Total Purchase',tp>0?fmtL(tp):'New party',C.blue],
      ['Avg / Month',st?.am?fmtL(st.am):'—',C.purple],
      ['Outstanding',od>0?fmtL(od):od<0?'Credit '+fmtL(-od):'Clear ✓',od>0?C.red:C.green],
      ['Grade',grade,gc[grade]],
    ].map(([l,v,col],i)=><div key={i} style={{background:col+'0D',borderRadius:7,padding:'7px 10px',border:`1px solid ${col}20`}}>
      <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.4,marginBottom:2}}>{l}</div>
      <div style={{fontSize:14,fontWeight:800,color:col,fontVariantNumeric:'tabular-nums'}}>{v}</div>
    </div>)}
  </div>;
}

// Renders the estimate as a PNG (via canvas, no library) so it can be shared
// as an image alongside the WhatsApp text — a plain wa.me link can only carry
// text, so an actual image needs either this + the Web Share API, or a manual
// download the user attaches themselves.
function buildInvoiceCanvas(sale,party){
  const W=680;
  const rows=[];
  rows.push({text:fmtD(sale.date)+'   ·   ESTIMATE',size:11,color:'#6B7280'});
  rows.push({text:party?.name||sale.partyName,size:17,color:'#111827',bold:true});
  if(sale.transport)rows.push({text:'Transport: '+sale.transport+(sale.bilty?'   |   LR: '+sale.bilty:''),size:12,color:'#4B5563'});
  if(sale.freightMode==='paid')rows.push({text:'Freight: '+fmt(sale.freightAmt)+' (paid by us)',size:12,color:'#4B5563'});
  else if(sale.freightMode==='topay')rows.push({text:'Freight: To Pay by party',size:12,color:'#4B5563'});
  rows.push({spacer:16});
  for(const l of (sale.items||[])){
    const name=(l.customerName||l.itemName)+(l.warranty&&l.warranty!=='no warranty'?'  ['+l.warranty+']':'');
    rows.push({cols:[name,l.qty+' × '+fmt(l.rate),fmt(l.amount)],size:13,color:'#111827'});
  }
  rows.push({spacer:10});
  rows.push({rule:true});
  rows.push({spacer:10});
  rows.push({cols:['TOTAL','',fmt(sale.total)],size:16,bold:true,color:'#2563EB'});
  const padX=28,padTop=34,rowH=25;
  let h=padTop;
  for(const r of rows)h+=r.spacer!=null?r.spacer:r.rule?9:rowH;
  h+=26;
  const scale=2;
  const canvas=document.createElement('canvas');
  canvas.width=W*scale;canvas.height=h*scale;
  const ctx=canvas.getContext('2d');
  ctx.scale(scale,scale);
  ctx.fillStyle='#FFFFFF';ctx.fillRect(0,0,W,h);
  let y=padTop;
  for(const r of rows){
    if(r.spacer!=null){y+=r.spacer;continue;}
    if(r.rule){ctx.strokeStyle='#E5E7EB';ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(padX,y);ctx.lineTo(W-padX,y);ctx.stroke();y+=9;continue;}
    ctx.font=(r.bold?'700 ':'400 ')+r.size+'px -apple-system,Segoe UI,sans-serif';
    ctx.fillStyle=r.color||'#111827';
    if(r.cols){
      ctx.textAlign='left';ctx.fillText(r.cols[0],padX,y,W-padX-220);
      ctx.textAlign='right';
      if(r.cols[1])ctx.fillText(r.cols[1],W-padX-110,y);
      ctx.fillText(r.cols[2],W-padX,y);
    } else {
      ctx.textAlign='left';ctx.fillText(r.text,padX,y);
    }
    y+=rowH;
  }
  return canvas;
}

// ── INVOICE MODAL ─────────────────────────────────────────────────────────────
function InvoiceModal({sale,party,onClose,viewOnly=false}){
  const[ph,setPh]=useState(party?.phone||'');
  const[phSaved,setPhSaved]=useState(!!party?.phone);
  const{toast,TC}=useToast();
  // Line items are editable right here — the customer-facing name is often
  // slightly different from what the item master has, and correcting it here
  // (a) fixes this invoice's text/image immediately and (b) is remembered for
  // *every* customer's invoices going forward, without overwriting item.customer
  // — that's the shared "system" name used by search and the item master.
  const[editableItems,setEditableItems]=useState(()=>(sale.items||[]).map(l=>({...l})));
  const updateName=(idx,val)=>setEditableItems(items=>items.map((l,i)=>i===idx?{...l,customerName:val}:l));
  const commitName=idx=>{
    const line=editableItems[idx];
    if(!line.itemId||!line.customerName)return;
    if(getItemCustomerName(line.itemId,null)!==line.customerName){
      setItemCustomerName(line.itemId,line.customerName);
      toast('Remembered "'+line.customerName+'" ✓');
    }
    if(sale.id){
      state.sales=state.sales.map(s=>s.id===sale.id?{...s,items:s.items.map(l=>l.itemId===line.itemId?{...l,customerName:line.customerName}:l)}:s);
      save.sales();
    }
  };
  const mergedItems=mergeBundledLines(editableItems);
  const effectiveSale={...sale,items:mergedItems};
  const due=()=>{if(!sale.creditDays)return'';const d=new Date(sale.date+'T00:00:00');d.setDate(d.getDate()+sale.creditDays);return d.toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'2-digit'});};
  const text=()=>applyTemplate('invoice',{
    date:fmtD(sale.date),partyName:sale.partyName,
    transport:sale.transport?'*Transport:* '+sale.transport+(sale.bilty?' | LR: '+sale.bilty:''):'',
    freight:sale.freightMode==='paid'?'*Freight:* Rs'+sale.freightAmt+' (paid by us)':sale.freightMode==='topay'?'*Freight:* To Pay by party':'',
    items:mergedItems.map(l=>'• '+(l.customerName||l.itemName)+' x '+l.qty+' pcs @ '+fmt(l.rate)+(l.warranty&&l.warranty!=='no warranty'?' ['+l.warranty+']':'')+' = *'+fmt(l.amount)+'*').join('\n'),
    total:fmt(sale.total),credit:sale.creditDays?'Credit '+sale.creditDays+'d | Due by '+due():'',
  });
  const savePh=()=>{if(!ph||ph.length<10)return toast('Enter 10-digit number','error');updatePartyPhone(party.id,ph);setPhSaved(true);toast('Saved ✓');};
  const[sharing,setSharing]=useState(false);
  const nameableIdx=editableItems.map((l,i)=>l.itemId&&!l.bundledJaali?i:-1).filter(i=>i>=0);
  const[step,setStep]=useState(!viewOnly&&nameableIdx.length>0?'names':'send');

  const sendText=()=>{window.open('https://wa.me/91'+ph+'?text='+encodeURIComponent(text()),'_blank');onClose();};
  // Pre-render the share image as soon as we reach the send step, so clicking
  // "Share Estimate" can call navigator.share() synchronously — canvas.toBlob()
  // is asynchronous, and doing it inside the click handler burns the browser's
  // user-activation window, which silently no-ops the WhatsApp share sheet.
  const[shareFile,setShareFile]=useState(null);
  useEffect(()=>{
    if(step!=='send'){setShareFile(null);return;}
    setSharing(true);
    let cancelled=false;
    buildInvoiceCanvas(effectiveSale,party).toBlob(blob=>{
      if(cancelled)return;
      setSharing(false);
      if(blob)setShareFile(new File([blob],'estimate-'+sale.id+'.png',{type:'image/png'}));
    },'image/png');
    return()=>{cancelled=true;};
  },[step]);
  const downloadImage=()=>{
    const url=URL.createObjectURL(shareFile);
    const a=document.createElement('a');
    a.href=url;a.download='estimate-'+sale.id+'.png';
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  // Desktop share sheets (macOS Chrome etc.) don't list WhatsApp as a target, and
  // using their "Copy" action pastes image+text as two separate items into
  // WhatsApp. So instead of navigator.share(), copy ONLY the image to the
  // clipboard (image/png ClipboardItem) and send the text as a separate,
  // always-available WhatsApp link — matching what actually pastes cleanly.
  const copyImage=async()=>{
    if(!shareFile)return;
    try{
      await navigator.clipboard.write([new ClipboardItem({'image/png':shareFile})]);
      toast('Image copied — paste it in WhatsApp ✓');
    }catch(e){
      downloadImage();
      toast('Copy not supported here — image downloaded instead','error');
    }
  };

  useEffect(()=>{
    const onDocKey=e=>{
      if(e.key==='Enter'&&e.target.tagName!=='INPUT'){
        e.preventDefault();
        if(step==='names')setStep('send');
        else if(phSaved&&ph)sendText();
      }else if(e.key==='Escape'){e.preventDefault();onClose();}
    };
    document.addEventListener('keydown',onDocKey);
    return()=>document.removeEventListener('keydown',onDocKey);
  },[step,phSaved,ph]);

  if(step==='names')return <Modal title="Confirm Customer-Facing Names" onClose={onClose} width={480}>
    <TC/>
    <div style={{fontSize:12,color:C.sub,marginBottom:12}}>How each item should read on estimates — edit any that don't match. Remembered for every customer going forward.</div>
    <div style={{border:`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden',marginBottom:16}}>
      {nameableIdx.map((idx,j)=><div key={idx} style={{display:'flex',alignItems:'center',gap:8,padding:'8px 10px',borderBottom:j<nameableIdx.length-1?`1px solid ${C.bdr}`:'none'}}>
        <span style={{fontSize:11,color:C.muted,flexShrink:0,maxWidth:110,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={editableItems[idx].itemName}>{editableItems[idx].itemName}</span>
        <input value={editableItems[idx].customerName||''} onChange={e=>updateName(idx,e.target.value)} onBlur={()=>commitName(idx)}
          onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();commitName(idx);setStep('send');}}}
          style={{flex:1,padding:'6px 8px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
      </div>)}
    </div>
    <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
      <Btn onClick={onClose} variant="secondary">Cancel (Esc)</Btn>
      <Btn onClick={()=>setStep('send')}>Confirm (Enter) →</Btn>
    </div>
  </Modal>;

  return <Modal title="Send Estimate" onClose={onClose} width={480}>
    <TC/>
    <div style={{background:'#F6F8FA',borderRadius:9,border:`1px solid ${C.bdr}`,padding:14,marginBottom:16,
      fontFamily:'monospace',fontSize:12,lineHeight:1.9,whiteSpace:'pre-wrap',color:C.txt,maxHeight:280,overflowY:'auto'}}>{text()}</div>
    {!phSaved&&<div style={{background:C.amberSoft,borderRadius:9,border:`1px solid ${C.amber}40`,padding:14,marginBottom:16}}>
      <div style={{fontSize:12,fontWeight:600,color:C.amberBright,marginBottom:8}}>No phone on file — add it to send via WhatsApp</div>
      <div style={{display:'flex',gap:8}}>
        <input value={ph} onChange={e=>setPh(e.target.value)} placeholder="WhatsApp number" autoComplete="off"
          style={{flex:1,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
        <Btn onClick={savePh} size="sm">Save</Btn>
      </div>
    </div>}
    <div style={{display:'flex',gap:8,marginBottom:6}}>
      <button onClick={copyImage} disabled={sharing||!shareFile} style={{flex:1,display:'flex',alignItems:'center',justifyContent:'center',gap:8,padding:'13px 14px',
          background:'#fff',color:C.txt,border:`1.5px solid ${C.bdr}`,borderRadius:9,cursor:(sharing||!shareFile)?'default':'pointer',fontWeight:700,fontSize:13,fontFamily:'inherit'}}>
        🖼️ {sharing?'Preparing…':'Copy Image'}
      </button>
      <button onClick={sendText} disabled={!phSaved||!ph} style={{flex:1,display:'flex',alignItems:'center',justifyContent:'center',gap:8,padding:'13px 14px',
          background:(phSaved&&ph)?'#25D366':C.bdr,color:'#fff',border:'none',borderRadius:9,cursor:(phSaved&&ph)?'pointer':'not-allowed',fontWeight:700,fontSize:13,fontFamily:'inherit'}}>
        📲 Send on WhatsApp (Enter)
      </button>
    </div>
    <div style={{fontSize:11,color:C.muted,textAlign:'center',marginBottom:10}}>Copy the image, open WhatsApp, paste it in the chat — the text sends separately as its own message.</div>
    <div style={{textAlign:'center',marginBottom:6}}>
      <button onClick={downloadImage} disabled={!shareFile} style={{background:'none',border:'none',cursor:shareFile?'pointer':'default',fontSize:11,color:C.muted,textDecoration:'underline',fontFamily:'inherit'}}>
        …or download the image instead
      </button>
    </div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginTop:10}}>
      <button onClick={()=>{navigator.clipboard?.writeText(text());toast('Copied');}}
        style={{padding:0,border:'none',background:'none',fontSize:11,color:C.muted,textDecoration:'underline',cursor:'pointer',fontFamily:'inherit'}}>Copy text</button>
      <Btn onClick={onClose} variant="secondary" size="sm">Not now (Esc)</Btn>
    </div>
  </Modal>;
}

// ── SALE ENTRY ────────────────────────────────────────────────────────────────
function SaleEntry({onSaved,mode='sale'}){
  const[date,setDate]=useState(today());
  const[credit,setCredit]=useState(30);
  const[partyId,setPartyId]=useState('');
  const[transport,setTransport]=useState('');
  const[bilty,setBilty]=useState('');
  const[fMode,setFMode]=useState('topay');
  const[fAmt,setFAmt]=useState('');
  const[fAddToInv,setFAddToInv]=useState(false);
  const[lines,setLines]=useState([]);
  const[input,setInput]=useState('');
  const[si,setSi]=useState(0);
  const[step,setStep]=useState('item');
  const[addPartyName,setAddPartyName]=useState(null);
  const[draft,setDraft]=useState({itemId:'',iName:'',cName:'',dName:'',peti:1,qty:'',rate:'',cat:'',warranty:'no warranty',withJaali:false});
  const[pq,setPq]=useState('');
  const[pIdx,setPIdx]=useState(0);
  const[saving,setSaving]=useState(false);
  const[done,setDone]=useState(null);
  const[trFocus,setTrFocus]=useState(false);
  const[trIdx,setTrIdx]=useState(0);
  const[ledgerParty,setLedgerParty]=useState(null);
  const{toast,TC}=useToast();
  const mobile=useIsMobile();
  const stepInputId='sale-step-input';

  useEffect(()=>{
    if(step!=='item'){setTimeout(()=>{const el=document.getElementById(stepInputId);if(el)el.focus();},30);}
  },[step,draft.iName]);

  useEffect(()=>{
    if(state.followups){const el=document.getElementById('isugg-'+si);if(el)el.scrollIntoView({block:'nearest'});}
  },[si]);

  const custs=useMemo(()=>state.parties.filter(p=>p.type==='customer').sort((a,b)=>a.name.localeCompare(b.name)),[state.parties.length]);
  const sel=useMemo(()=>custs.find(p=>p.id===partyId),[custs,partyId]);
  const psugg=useMemo(()=>pq.length<1?[]:custs.filter(p=>p.name.toLowerCase().includes(pq.toLowerCase())||p.city?.toLowerCase().includes(pq.toLowerCase())).slice(0,8),[custs,pq]);
  const freightLine=useMemo(()=>fMode==='paid'&&fAddToInv&&Number(fAmt)>0?[{id:'__freight',itemName:'Freight',customerName:'Freight',peti:1,qty:1,rate:Number(fAmt)}]:[]  ,[fMode,fAddToInv,fAmt]);
  const total=useMemo(()=>[...lines,...freightLine].reduce((s,l)=>s+(l.qty*l.rate),0),[lines,freightLine]);
  const prevItems=useMemo(()=>sel?getPrevItems(sel.name):[]  ,[sel]);
  const prevItemIds=useMemo(()=>new Set(prevItems.map(p=>p.item.id)),[prevItems]);
  const expectedMap=useMemo(()=>getExpectedStockMap(),[state.purchases.length]);
  const itemFreq=useMemo(()=>{
    const map={};
    for(const m of getAllItemMetrics())map[m.item.id]=m.totalSoldQty;
    return map;
  },[state.sales.length]);
  const allTransporterNames=useMemo(()=>getAllTransporterNames(),[state.sales.length]);
  const cityTopTransporters=useMemo(()=>sel?getTopTransportersForCity(sel.city):[],[sel]);
  const trSuggestions=useMemo(()=>{
    const q=transport.trim().toLowerCase();
    const pool=q?allTransporterNames.filter(n=>n.toLowerCase().includes(q)):[...new Set([...cityTopTransporters,...allTransporterNames])];
    return pool.slice(0,6);
  },[transport,allTransporterNames,cityTopTransporters]);

  useEffect(()=>{
    if(sel){
      setCredit(sel.credit||30);
      const remembered=state.partyTransport[sel.id];
      if(remembered)setTransport(remembered);
      else{const top=getTopTransportersForCity(sel.city)[0];if(top)setTransport(top);}
      setTimeout(()=>{const el=document.getElementById('item-search-input');if(el)el.focus();},50);
    }
  },[sel]);

  const isugg=useMemo(()=>{
    const q=input.trim().toLowerCase();if(!q||step!=='item')return[];
    const words=q.split(/[\s,]+/).filter(Boolean);
    return state.items.filter(it=>it.active!==false).map(it=>{
      const cn=(it.customer||it.internal||it.name||'').toLowerCase();
      const intn=(it.internal||it.name||'').toLowerCase();
      const hay=[cn,intn,(it.cat||'').toLowerCase(),(it.supplier||'').toLowerCase()].join(' ');
      let score=0;
      for(const w of words){if(cn.includes(w))score+=10;if(intn.includes(w))score+=8;if(hay.includes(w))score+=4;if(cn.startsWith(w))score+=10;}
      if(cn.includes(q))score+=15;
      if(prevItemIds.has(it.id))score+=20;
      score+=Math.min(15,(itemFreq[it.id]||0)/20);
      return{it,score};
    }).filter(x=>x.score>0).sort((a,b)=>b.score-a.score).slice(0,10).map(x=>x.it);
  },[input,step,prevItemIds,itemFreq]);

  const pickItem=it=>{
    const rate=sel?getBestRate(sel.name,it.internal||'',it.customer||'',it.sale,it.id):it.sale;
    const cName=getItemCustomerName(it.id,it.customer||it.internal||'');
    setDraft({itemId:it.id,iName:it.internal||it.name||'',cName,dName:it.dispatchName||it.internal||'',peti:it.peti,qty:'',rate,cat:it.cat,warranty:'no warranty',withJaali:false});
    setStep('qty');setInput('');setSi(0);
  };

  const commit=()=>{
    const qty=parseQty(draft.qty,draft.peti);
    if(isNaN(qty)||qty<=0)return toast('Invalid qty','error');
    const bundleJaali=draft.cat==='Infrared'&&draft.withJaali;
    const mainName=bundleJaali?'Infrared +Jaali':draft.cName;
    const mainDName=bundleJaali?'Infrared +Jaali':draft.dName;
    const newLines=[{id:uid('L'),itemId:draft.itemId,itemName:draft.iName,customerName:mainName,dispatchName:mainDName,peti:draft.peti,qty,rate:Number(draft.rate)||0,warranty:DAMAGE_TRACKED_CATS.has(draft.cat)?draft.warranty:undefined}];
    if(bundleJaali){
      const jaali=state.items.find(x=>x.active!==false&&(x.internal||'').toLowerCase()==='jaali');
      if(jaali){
        // No default rate — the user fills this in on the line-items table.
        newLines.push({id:newLines[0].id+'J',itemId:jaali.id,itemName:jaali.internal||jaali.name||'',customerName:jaali.customer||jaali.internal||'',dispatchName:jaali.dispatchName||jaali.internal||'',peti:jaali.peti,qty,rate:0,bundledJaali:true});
      }
    }
    setLines(p=>[...p,...newLines]);
    setDraft({itemId:'',iName:'',cName:'',dName:'',peti:1,qty:'',rate:'',cat:'',warranty:'no warranty',withJaali:false});setStep('item');setInput('');setSi(0);
    setTimeout(()=>document.getElementById('item-search-input')?.focus(),30);
  };

  const onKey=e=>{
    if(step==='item'){
      if(e.key==='ArrowDown'){e.preventDefault();setSi(i=>Math.min(i+1,(isugg.length||1)-1));}
      if(e.key==='ArrowUp')  {e.preventDefault();setSi(i=>Math.max(i-1,0));}
      if(e.key==='Enter'&&isugg.length){e.preventDefault();pickItem(isugg[si]||isugg[0]);}
      else if(e.key==='Enter'&&!input.trim()){e.preventDefault();document.getElementById('sale-transport-input')?.focus();}
      if(e.key==='Escape'){setInput('');setSi(0);}
    }else if(step==='qty'){
      if(e.key==='Enter'){e.preventDefault();if(!draft.qty)return;const q=parseQty(draft.qty,draft.peti);if(isNaN(q))return toast('Invalid qty','error');setDraft(d=>({...d,qty:String(q)}));setStep('rate');}
      if(e.key==='Escape'){setStep('item');setInput('');}
    }else if(step==='rate'){
      if(e.key==='Enter'){e.preventDefault();if(DAMAGE_TRACKED_CATS.has(draft.cat))setStep('warranty');else commit();}
      if(e.key==='Escape'){setStep('qty');}
    }else if(step==='warranty'){
      if(e.key==='Enter'){e.preventDefault();commit();}
      if(e.key==='Escape'){setStep('rate');}
    }
  };

  const editLine=(i,k,v)=>setLines(p=>p.map((l,j)=>j!==i?l:{...l,[k]:Number(v)||0}));
  const resetForm=()=>{
    setPartyId('');setPq('');setLines([]);setStep('item');setInput('');setTransport('');setBilty('');
    setFMode('topay');setFAmt('');setFAddToInv(false);
  };
  const doSave=async()=>{
    if(!partyId)return toast('Select a party','error');
    if(!lines.length)return toast('Add at least one item','error');
    setSaving(true);
    const allItems=[...lines,...freightLine].map(l=>({itemId:l.itemId,itemName:l.itemName,customerName:l.customerName,dispatchName:l.dispatchName,qty:l.qty,rate:l.rate,amount:l.qty*l.rate,warranty:l.warranty,bundledJaali:l.bundledJaali}));
    if(mode==='order'){
      const order={id:uid('ORD'),date,partyId,partyName:sel?.name||'',transport,bilty,freightMode:fMode,freightAmt:Number(fAmt)||0,freightAddToInv:fAddToInv,creditDays:credit,items:allItems,total,status:'pending',createdAt:new Date().toISOString()};
      state.orders=[order,...state.orders];save.orders();
      if(transport&&sel)recordCityTransport(sel.city,transport);
      setSaving(false);
      resetForm();
      onSaved?.(order);
      return;
    }
    const id=uid('SAL');
    const sale={id,date,partyId,partyName:sel?.name||'',transport,bilty,freightMode:fMode,freightAmt:Number(fAmt)||0,freightAddToInv:fAddToInv,creditDays:credit,items:allItems,total,createdAt:new Date().toISOString()};
    state.sales=[sale,...state.sales];save.sales();
    if(transport&&sel)recordCityTransport(sel.city,transport);
    lines.forEach(l=>adjustStock(l.itemId,-l.qty));
    setSaving(false);setDone(sale);
  };

  const SC={item:C.blue,qty:C.green,rate:C.amber,warranty:C.purple};
  const SL={item:'Type anything — name, category, model…',qty:'Qty, "1 peti" or "1p" = '+draft.peti+'pc',rate:'Rate Rs/pc — Enter to add',warranty:'Select warranty — Enter to add'};

  if(done)return <InvoiceModal sale={done} party={sel} onClose={()=>{
    setDone(null);setPartyId('');setTransport('');setBilty('');setFMode('topay');setFAmt('');setFAddToInv(false);
    setPq('');setLines([]);setStep('item');setInput('');onSaved?.();
    setTimeout(()=>document.getElementById('sale-date-input')?.focus(),30);
  }}/>;

  return <div>
    <TC/>
    {addPartyName!==null&&<QuickAddPartyModal defaultName={addPartyName} onClose={()=>setAddPartyName(null)} onAdded={p=>{setPartyId(p.id);setPq('');setAddPartyName(null);}}/>}
    {ledgerParty&&<PartyLedger party={ledgerParty} onClose={()=>setLedgerParty(null)}/>}
    {/* Date + Credit */}
    <div style={{display:'flex',gap:10,marginBottom:10,flexWrap:'wrap'}}>
      <div style={{flex:mobile?'1 1 140px':'0 0 155px'}}><Inp id="sale-date-input" label="Date" value={date} onChange={setDate} type="date" required
        onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();document.getElementById('sale-credit-input')?.focus();}}}/></div>
      <div style={{flex:mobile?'1 1 120px':'0 0 125px'}}><Inp id="sale-credit-input" label="Credit Days" value={credit} onChange={v=>setCredit(Number(v)||30)} type="number" hint={sel?'Default: '+sel.credit+'d':''}
        onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();document.getElementById(sel?'item-search-input':'party-search-input')?.focus();}}}/></div>
    </div>

    {/* Party */}
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'12px 16px',marginBottom:10}}>
      <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:4}}>Party <span style={{color:C.red}}>*</span></label>
      {sel
        ?<div>
          <div style={{display:'flex',alignItems:'center',gap:10,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
            <div style={{flex:1}}>
              <span onClick={()=>setLedgerParty(sel)} style={{fontWeight:700,color:C.blue,cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted'}}>{sel.name}</span>
              <span style={{fontSize:11,color:C.sub,marginLeft:8}}>{sel.city} · {sel.route} · {credit}d credit</span>
            </div>
            <Btn onClick={()=>{setPartyId('');setPq('');setCredit(30);setTransport('');}} variant="ghost" size="sm">Change</Btn>
          </div>
          <PartyInfoCard party={sel}/>
        </div>
        :<div style={{position:'relative'}}>
          <input id="party-search-input" autoFocus value={pq} autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
            onChange={e=>{setPq(e.target.value);setPIdx(0);}}
            onKeyDown={e=>{
              if(e.key==='ArrowDown'){e.preventDefault();setPIdx(i=>Math.min(i+1,(psugg.length||1)-1));}
              if(e.key==='ArrowUp')  {e.preventDefault();setPIdx(i=>Math.max(i-1,0));}
              if(e.key==='Enter'&&psugg.length){e.preventDefault();setPartyId(psugg[pIdx]?.id||psugg[0].id);setPq('');}
              else if(e.key==='Enter'&&pq.length>=2){e.preventDefault();setAddPartyName(pq);}
              if(e.key==='Escape'){setPq('');setPIdx(0);}
            }}
            placeholder="Start typing party name or city…"
            style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
            onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
          {(psugg.length>0||pq.length>=2)&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
            {psugg.map((p,pi)=><div key={p.id} onMouseDown={()=>{setPartyId(p.id);setPq('');}} onMouseEnter={()=>setPIdx(pi)}
              style={{padding:'9px 14px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`,background:pi===pIdx?C.blueSoft:'transparent'}}>
              <span style={{fontWeight:600,color:C.txt}}>{p.name}</span>
              <span style={{fontSize:11,color:C.sub,marginLeft:8}}>{p.city} · {p.route}</span>
              {(()=>{const _od=getCustomerBal(p.id);return _od!==0&&<span style={{fontSize:11,color:_od>0?C.red:C.green,marginLeft:8,fontWeight:600}}>{_od>0?'OD '+fmt(_od):'Credit '+fmt(-_od)}</span>;})()}
            </div>)}
            {pq.length>=2&&psugg.length===0&&<div onMouseDown={()=>setAddPartyName(pq)}
              style={{padding:'9px 14px',cursor:'pointer',color:C.blue,fontWeight:600,fontSize:13,background:C.blueSoft}}>
              + Add "{pq}" as new party
            </div>}
          </div>}
        </div>}
    </div>

    {/* Item entry + lines */}
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',marginBottom:10}}>
      {/* Previously bought */}
      {prevItems.length>0&&step==='item'&&<div style={{padding:'10px 14px',borderBottom:`1px solid ${C.bdr}`,background:'#FAFBFC'}}>
        <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:7}}>Previously bought ({prevItems.length})</div>
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(130px,1fr))',gap:6}}>
          {prevItems.map(({item,rate},i)=>(
            <button key={i} onMouseDown={()=>pickItem(item)} style={{display:'flex',flexDirection:'column',alignItems:'center',padding:'6px 10px',borderRadius:7,cursor:'pointer',fontFamily:'inherit',border:`1.5px solid ${C.bdr}`,background:C.surf,width:'100%'}}
              onMouseEnter={e=>{e.currentTarget.style.borderColor=C.blue;e.currentTarget.style.background=C.blueSoft;}}
              onMouseLeave={e=>{e.currentTarget.style.borderColor=C.bdr;e.currentTarget.style.background=C.surf;}}>
              <span style={{fontSize:11,fontWeight:600,color:C.txt,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',width:'100%',textAlign:'center'}}>{item.customer||item.internal}</span>
              <span style={{fontSize:11,fontWeight:700,color:C.blue,textAlign:'center'}}>₹{rate}/pc</span>
            </button>
          ))}
        </div>
      </div>}

      {/* Step input */}
      <div style={{padding:'10px 14px',background:step==='item'?C.blueSoft:step==='qty'?C.greenSoft:step==='rate'?C.amberSoft:C.purpleSoft,borderBottom:`1px solid ${C.bdr}`,position:'relative'}}>
        <div style={{display:'flex',gap:8,alignItems:'center'}}>
          {step!=='item'&&<span style={{fontSize:12,fontWeight:600,color:C.txt,whiteSpace:'nowrap',maxWidth:160,overflow:'hidden',textOverflow:'ellipsis'}}>
            {step==='rate'?`${draft.cName||draft.iName} × ${draft.qty}pc`:step==='warranty'?`${draft.cName||draft.iName} × ${draft.qty}pc @ ₹${draft.rate}`:draft.cName||draft.iName}
          </span>}
          {step==='warranty'
            ?<select id={stepInputId} value={draft.warranty} onChange={e=>setDraft(d=>({...d,warranty:e.target.value}))} onKeyDown={onKey}
              style={{flex:1,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${SC.warranty}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',fontWeight:500,background:C.surf}}>
              {WARRANTY_OPTIONS.map(w=><option key={w} value={w}>{w}</option>)}
            </select>
            :<input id={step==='item'?'item-search-input':stepInputId} key={step+draft.iName}
              autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
              value={step==='item'?input:step==='qty'?draft.qty:String(draft.rate||'')}
              onChange={e=>{const v=e.target.value;if(step==='item'){setInput(v);setSi(0);}else if(step==='qty')setDraft(d=>({...d,qty:v}));else setDraft(d=>({...d,rate:v}));}}
              onKeyDown={onKey} placeholder={SL[step]}
              style={{flex:1,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${SC[step]}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',fontWeight:500}}/>}
          {step!=='item'&&<><Btn onClick={commit} size="sm" variant="success">Add</Btn><Btn onClick={()=>{setStep('item');setInput('');}} size="sm" variant="ghost">✕</Btn></>}
        </div>
        {step==='warranty'&&draft.cat==='Infrared'&&<label style={{display:'flex',alignItems:'center',gap:7,marginTop:9,fontSize:12,color:C.txt,cursor:'pointer'}}>
          <input type="checkbox" checked={draft.withJaali} onChange={e=>setDraft(d=>({...d,withJaali:e.target.checked}))}
            onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();commit();}}}/>
          Include Jaali (net) — same qty
        </label>}
        {/* Suggestions */}
        {step==='item'&&isugg.length>0&&<div style={{position:'absolute',top:'calc(100% - 2px)',left:14,right:14,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden',maxHeight:180,display:'flex',flexDirection:'column'}}>
          <div style={{overflowY:'auto',flex:1}}>
            {isugg.map((it,i)=>{
              const r=sel?getBestRate(sel.name,it.internal||'',it.customer||'',it.sale,it.id):it.sale;
              const expected=expectedMap[it.id]||0;
              return <div key={it.id} id={'isugg-'+i} onMouseDown={()=>{setSi(i);pickItem(it);}} onMouseEnter={()=>setSi(i)}
                style={{display:'flex',alignItems:'center',padding:'8px 14px',cursor:'pointer',background:i===si?C.blueSoft:'transparent',borderBottom:i<isugg.length-1?`1px solid ${C.bdr}`:'none'}}>
                <div style={{flex:1}}>
                  <span style={{fontWeight:600,color:C.txt,fontSize:13}}>{it.internal||it.name}</span>
                  {it.customer&&it.customer!==it.internal&&<span style={{fontSize:10,color:C.muted,marginLeft:6}}>{it.customer}</span>}
                  <span style={{fontSize:11,color:C.sub,marginLeft:6}}>{it.cat} · {it.peti}pc/peti</span>
                  <div style={{fontSize:11,marginTop:1}}>
                    <span style={{color:(it.stock||0)<=0?C.red:C.sub,fontWeight:(it.stock||0)<=0?700:400}}>stock {it.stock||0}</span>
                    {expected>0&&<span style={{color:C.amber,fontWeight:600,marginLeft:6}}>+{expected} expected</span>}
                  </div>
                </div>
                <div style={{textAlign:'right',flexShrink:0}}>
                  <span style={{fontSize:13,fontWeight:700,color:r!==it.sale?C.amber:C.blue}}>₹{r}/pc</span>
                  {r!==it.sale&&<div style={{fontSize:10,color:C.muted}}>std ₹{it.sale}</div>}
                </div>
              </div>;
            })}
          </div>
          <div style={{padding:'4px 14px',fontSize:10,color:C.muted,background:'#F6F8FA',flexShrink:0}}>↑↓ · Enter select · Esc cancel</div>
        </div>}
        {step==='qty'&&draft.qty&&!isNaN(parseQty(draft.qty,draft.peti))&&<div style={{marginTop:4,fontSize:11,color:C.green}}>→ {parseQty(draft.qty,draft.peti)} pcs</div>}
      </div>

      {/* Lines */}
      <CardTable mobile={mobile} rows={[...lines,...freightLine]} keyFn={(l,i)=>l.id||i}
        emptyLabel={sel?'Type item or click chip above':'Select a party first'}
        footer={(lines.length>0||freightLine.length>0)?
          <div style={{display:'flex',justifyContent:'space-between',padding:'9px 12px',background:'#F6F8FA'}}>
            <span style={{fontWeight:700,color:C.txt}}>TOTAL</span>
            <span style={{fontSize:16,fontWeight:800,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(total)}</span>
          </div>:null}
        columns={[
          {key:'idx',label:'#',hideOnMobile:true,render:(l,i)=><span style={{color:C.muted,fontSize:11}}>{i+1}</span>},
          {key:'item',label:'Item',render:l=><div>
            <div style={{fontWeight:600,color:l.id==='__freight'?C.amber:C.txt}}>{l.customerName||l.itemName}</div>
            {l.customerName&&l.customerName!==l.itemName&&<div style={{fontSize:10,color:C.muted}}>{l.itemName}</div>}
            {l.warranty&&l.warranty!=='no warranty'&&<div style={{fontSize:10,color:C.purple,fontWeight:600}}>🛡 {l.warranty}</div>}
          </div>},
          {key:'qty',label:'Qty',align:'right',render:(l,i)=>
            l.id!=='__freight'?<input value={l.qty} onChange={e=>editLine(i,'qty',e.target.value)} type="number" style={{width:60,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/>:<span style={{color:C.amber,fontSize:12}}>1</span>},
          {key:'rate',label:'Rate ₹',align:'right',render:(l,i)=>
            l.id!=='__freight'?<input value={l.rate} onChange={e=>editLine(i,'rate',e.target.value)} type="number" style={{width:75,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/>:<span style={{color:C.amber,fontSize:12}}>{fmt(l.rate)}</span>},
          {key:'amount',label:'Amount',align:'right',render:l=><span style={{fontWeight:700,color:l.id==='__freight'?C.amber:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(l.qty*l.rate)}</span>},
          {key:'actions',label:'',cardLabel:'',render:(l,i)=>l.id!=='__freight'?
            <Btn onClick={()=>setLines(p=>p.filter((_,j)=>j!==i))} variant="ghost" size="sm">Remove</Btn>:null},
        ]}/>
    </div>

    {/* Transport + Freight */}
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'12px 16px',marginBottom:10}}>
      <Grid2>
        <div style={{position:'relative'}}>
          <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:4}}>Transport</label>
          <input id="sale-transport-input" value={transport} autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
            onChange={e=>{setTransport(e.target.value);setTrIdx(0);if(sel){state.partyTransport[sel.id]=e.target.value;save.partyTransport();}}}
            onFocus={e=>{setTrFocus(true);e.target.style.borderColor=C.blue;}}
            onBlur={e=>{setTrFocus(false);e.target.style.borderColor=C.bdr;}}
            onKeyDown={e=>{
              if(e.key==='ArrowDown'){e.preventDefault();setTrIdx(i=>Math.min(i+1,(trSuggestions.length||1)-1));}
              if(e.key==='ArrowUp'){e.preventDefault();setTrIdx(i=>Math.max(i-1,0));}
              if(e.key==='Enter'){
                e.preventDefault();
                if(trSuggestions.length){
                  const name=trSuggestions[trIdx]||trSuggestions[0];
                  setTransport(name);if(sel){state.partyTransport[sel.id]=name;save.partyTransport();}
                }
                setTrFocus(false);
                document.getElementById('sale-bilty-input')?.focus();
              }
              if(e.key==='Escape')setTrFocus(false);
            }}
            placeholder="Garage / transporter name"
            style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}/>
          {trFocus&&trSuggestions.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
            {trSuggestions.map((name,i)=><div key={i}
              onMouseDown={()=>{setTransport(name);setTrIdx(i);if(sel){state.partyTransport[sel.id]=name;save.partyTransport();}}}
              onMouseEnter={()=>setTrIdx(i)}
              style={{padding:'8px 14px',cursor:'pointer',fontSize:13,borderBottom:i<trSuggestions.length-1?`1px solid ${C.bdr}`:'none',background:i===trIdx?C.blueSoft:'transparent'}}>
              <span style={{fontWeight:600,color:C.txt}}>{name}</span>
              {cityTopTransporters.includes(name)&&sel&&<span style={{fontSize:10,color:C.blue,marginLeft:6}}>· used for {sel.city}</span>}
            </div>)}
          </div>}
        </div>
        <Inp id="sale-bilty-input" label="Bilty / LR No." value={bilty} onChange={setBilty} placeholder="LR number" autoComplete="off"
          onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();const next=fMode==='paid'?document.getElementById('sale-freight-amt-input'):null;if(next)next.focus();else doSave();}}}/>
      </Grid2>
      <div style={{marginTop:10}}>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:6}}>Freight</label>
        <div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
          {[['topay','To Pay by party'],['','Not applicable'],['paid','Paid by us']].map(([val,lbl])=>(
            <button key={val} onClick={()=>{setFMode(val);if(val!=='paid')setFAddToInv(false);}} style={{padding:'5px 13px',borderRadius:20,fontSize:12,fontWeight:600,cursor:'pointer',fontFamily:'inherit',border:`1.5px solid ${fMode===val?C.blue:C.bdr}`,background:fMode===val?C.blueSoft:'transparent',color:fMode===val?C.blue:C.sub}}>{lbl}</button>
          ))}
          {fMode==='paid'&&<>
            <input id="sale-freight-amt-input" value={fAmt} onChange={e=>setFAmt(e.target.value)} type="number" placeholder="Rs amount" autoComplete="off"
              onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();doSave();}}}
              style={{width:90,padding:'5px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:12,outline:'none',fontFamily:'inherit'}}/>
            <label style={{display:'flex',alignItems:'center',gap:5,fontSize:12,color:C.txt,cursor:'pointer'}}>
              <input type="checkbox" checked={fAddToInv} onChange={e=>setFAddToInv(e.target.checked)}/>Add to invoice
            </label>
          </>}
        </div>
      </div>
    </div>

    {/* Save */}
    <div style={{display:'flex',flexDirection:mobile?'column-reverse':'row',justifyContent:'flex-end',gap:8}}>
      <Btn onClick={resetForm} variant="secondary" size="sm">Clear</Btn>
      <Btn onClick={doSave} disabled={saving||!partyId||!lines.length} size="md" icon={mode==='order'?'📲':'💾'}>{saving?'Saving…':(mode==='order'?'Send to Dispatch · ':'Save · ')+fmt(total)}</Btn>
    </div>
  </div>;
}

// ── DISPATCH PREVIEW ──────────────────────────────────────────────────────────
function DispatchPreviewModal({order,onClose}){
  const{toast,TC}=useToast();
  const[phone,setPhone]=useState(state.dispatchPhone||'');
  const[editableItems,setEditableItems]=useState(()=>(order.items||[]).map(l=>({...l})));
  const updateName=(idx,val)=>setEditableItems(items=>items.map((l,i)=>i===idx?{...l,dispatchName:val}:l));
  const commitName=idx=>{
    const line=editableItems[idx];
    if(!line.itemId||!line.dispatchName)return;
    const it=state.items.find(x=>x.id===line.itemId);
    if(it&&it.dispatchName!==line.dispatchName){
      state.items=state.items.map(x=>x.id===line.itemId?{...x,dispatchName:line.dispatchName}:x);
      save.items();
      toast('Remembered "'+line.dispatchName+'" for dispatch ✓');
    }
    state.orders=state.orders.map(o=>o.id===order.id?{...o,items:o.items.map(l=>l.itemId===line.itemId?{...l,dispatchName:line.dispatchName}:l)}:o);
    save.orders();
  };
  const party=state.parties.find(p=>p.id===order.partyId);
  const totalQty=editableItems.reduce((s,l)=>s+(l.qty||0),0);
  const totalAmount=editableItems.reduce((s,l)=>s+(l.amount??(l.qty*l.rate)),0);
  const mergedItems=mergeBundledLines(editableItems);
  const text=()=>applyTemplate('dispatch',{
    date:fmtD(order.date),partyName:order.partyName,city:party?.city||'',
    items:mergedItems.map(l=>'• '+(l.dispatchName||l.itemName)+' x '+l.qty+' pcs @ '+fmt(l.rate)+' = '+fmt(l.amount??(l.qty*l.rate))).join('\n'),
    totalQty,totalAmount:fmt(totalAmount),
    transport:order.transport?'*Transport:* '+order.transport+(order.bilty?' | LR: '+order.bilty:''):'',
  });
  const savePhone=()=>{
    if(!phone||phone.length<10)return toast('Enter 10-digit number','error');
    state.dispatchPhone=phone;save.dispatchPhone();toast('Saved ✓');
  };
  const nameableIdx=editableItems.map((l,i)=>l.itemId&&!l.bundledJaali?i:-1).filter(i=>i>=0);
  return <Modal title="Send to Dispatch" onClose={onClose} width={500}>
    <TC/>
    {nameableIdx.length>0&&<div style={{marginBottom:14}}>
      <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:6}}>
        Dispatch names <span style={{fontWeight:400,color:C.muted,textTransform:'none',letterSpacing:0}}>— edit to correct, remembered for next time</span>
      </div>
      <div style={{border:`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden'}}>
        {nameableIdx.map((idx,j)=><div key={idx} style={{display:'flex',alignItems:'center',gap:8,padding:'6px 10px',borderBottom:j<nameableIdx.length-1?`1px solid ${C.bdr}`:'none'}}>
          <span style={{fontSize:11,color:C.muted,flexShrink:0,maxWidth:110,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={editableItems[idx].itemName}>{editableItems[idx].itemName}</span>
          <input value={editableItems[idx].dispatchName||''} onChange={e=>updateName(idx,e.target.value)} onBlur={()=>commitName(idx)}
            style={{flex:1,padding:'4px 8px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,outline:'none',fontFamily:'inherit'}}/>
        </div>)}
      </div>
    </div>}
    <div style={{background:'#F6F8FA',borderRadius:9,border:`1px solid ${C.bdr}`,padding:14,marginBottom:16,
      fontFamily:'monospace',fontSize:12,lineHeight:1.9,whiteSpace:'pre-wrap',color:C.txt,maxHeight:260,overflowY:'auto'}}>{text()}</div>
    {state.dispatchPhone
      ?<a href={'https://wa.me/91'+state.dispatchPhone+'?text='+encodeURIComponent(text())} target="_blank" rel="noreferrer" onClick={onClose}
        style={{display:'flex',alignItems:'center',justifyContent:'center',gap:10,padding:'13px 20px',
          background:'#25D366',color:'#fff',borderRadius:9,textDecoration:'none',fontWeight:700,fontSize:14,marginBottom:16}}>📲 Send to Dispatch · {state.dispatchPhone}</a>
      :<div style={{background:C.amberSoft,borderRadius:9,border:`1px solid ${C.amber}40`,padding:14,marginBottom:16}}>
        <div style={{fontSize:12,fontWeight:600,color:C.amberBright,marginBottom:8}}>No dispatch number set — add once, reused every time</div>
        <div style={{display:'flex',gap:8}}>
          <input value={phone} onChange={e=>setPhone(e.target.value)} placeholder="Dispatch WhatsApp number" autoComplete="off"
            style={{flex:1,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
          <Btn onClick={savePhone} size="sm">Save</Btn>
        </div>
      </div>}
    <div style={{display:'flex',gap:8,justifyContent:'space-between',alignItems:'center'}}>
      <button onClick={()=>{navigator.clipboard?.writeText(text());toast('Copied');}}
        style={{padding:'6px 14px',borderRadius:7,border:`1px solid ${C.bdr}`,background:C.surf,fontSize:12,cursor:'pointer',fontFamily:'inherit',color:C.sub}}>📋 Copy</button>
      <Btn onClick={onClose} variant="secondary">Close</Btn>
    </div>
  </Modal>;
}

// ── REVIEW ORDER → INVOICE ────────────────────────────────────────────────────
function ReviewOrderModal({order,onClose,onSaved}){
  const[lines,setLines]=useState(()=>(order.items||[]).map(l=>({...l})));
  const{toast,TC}=useToast();
  const editLine=(i,k,v)=>setLines(p=>p.map((l,j)=>j!==i?l:{...l,[k]:Number(v)||0}));
  const total=lines.reduce((s,l)=>s+(l.qty*l.rate),0);
  const doCreateInvoice=()=>{
    if(!lines.length)return toast('No items','error');
    const id=uid('SAL');
    const allItems=lines.map(l=>({itemId:l.itemId,itemName:l.itemName,customerName:l.customerName,dispatchName:l.dispatchName,qty:l.qty,rate:l.rate,amount:l.qty*l.rate,warranty:l.warranty,bundledJaali:l.bundledJaali}));
    const sale={id,date:today(),partyId:order.partyId,partyName:order.partyName,transport:order.transport,bilty:order.bilty,
      freightMode:order.freightMode,freightAmt:order.freightAmt,freightAddToInv:order.freightAddToInv,creditDays:order.creditDays,
      items:allItems,total,orderId:order.id,createdAt:new Date().toISOString()};
    state.sales=[sale,...state.sales];save.sales();
    lines.forEach(l=>adjustStock(l.itemId,-l.qty));
    state.orders=state.orders.map(o=>o.id===order.id?{...o,status:'invoiced'}:o);
    save.orders();
    onSaved?.(sale);
  };
  useEffect(()=>{
    const onKey=e=>{
      if(e.target.tagName==='INPUT')return;
      if(e.key==='Enter'){e.preventDefault();doCreateInvoice();}
      if(e.key==='Escape'){e.preventDefault();onClose();}
    };
    document.addEventListener('keydown',onKey);
    return()=>document.removeEventListener('keydown',onKey);
  },[lines]);
  return <Modal title="Review Before Invoice" onClose={onClose} width={560}>
    <TC/>
    <div style={{fontSize:12,color:C.sub,marginBottom:12}}>Confirm what was actually dispatched — adjust qty/rate if it differs from the order, then create the invoice.</div>
    <table style={{width:'100%',borderCollapse:'collapse',fontSize:13,marginBottom:14}}>
      <thead><tr style={{background:'#F6F8FA'}}>
        {['Item','Qty','Rate ₹','Amount'].map((h,i)=><th key={i} style={{padding:'7px 12px',textAlign:i>0?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}
      </tr></thead>
      <tbody>
        {lines.map((l,i)=><tr key={l.id||i} style={{borderBottom:`1px solid ${C.bdr}`}}>
          <td style={{padding:'7px 12px',fontWeight:600,color:C.txt}}>{l.customerName||l.itemName}</td>
          <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.qty} onChange={e=>editLine(i,'qty',e.target.value)} type="number" style={{width:60,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
          <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.rate} onChange={e=>editLine(i,'rate',e.target.value)} type="number" style={{width:75,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
          <td style={{padding:'7px 12px',textAlign:'right',fontWeight:700,color:C.blue}}>{fmt(l.qty*l.rate)}</td>
        </tr>)}
      </tbody>
      <tfoot><tr style={{background:'#F6F8FA'}}>
        <td colSpan={3} style={{padding:'9px 12px',fontWeight:700,textAlign:'right',color:C.txt}}>TOTAL</td>
        <td style={{padding:'9px 12px',textAlign:'right',fontSize:16,fontWeight:800,color:C.blue}}>{fmt(total)}</td>
      </tr></tfoot>
    </table>
    <div style={{fontSize:11,color:C.muted,marginBottom:10}}>Press Enter to create the invoice, Esc to cancel.</div>
    <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
      <Btn onClick={onClose} variant="secondary">Cancel (Esc)</Btn>
      <Btn onClick={doCreateInvoice} icon="🧾">Create Invoice (Enter)</Btn>
    </div>
  </Modal>;
}

// ── NEW ORDER MODULE ──────────────────────────────────────────────────────────
function OrdersModule(){
  const[dispatchOrder,setDispatchOrder]=useState(null);
  const[reviewOrder,setReviewOrder]=useState(null);
  const[doneSale,setDoneSale]=useState(null);
  const[_fv,_setFv]=useState(0); const forceUpdate=()=>_setFv(v=>v+1);

  const pending=useMemo(()=>state.orders.filter(o=>o.status!=='invoiced').sort((a,b)=>b.createdAt.localeCompare(a.createdAt)),[state.orders.length,_fv]);

  const openReview=o=>{
    if(o.status!=='dispatched'){state.orders=state.orders.map(x=>x.id===o.id?{...x,status:'dispatched',dispatchedAt:new Date().toISOString()}:x);save.orders();forceUpdate();}
    setReviewOrder(o);
  };

  return <div>
    <SectionHead title="Sales Order" sub="Punch the order, send it to dispatch, then create the invoice once it's out the door"/>
    <SaleEntry mode="order" onSaved={order=>{forceUpdate();setDispatchOrder(order);}}/>

    {pending.length>0&&<div>
      <h3 style={{fontSize:14,fontWeight:700,color:C.txt,margin:'0 0 10px'}}>Sales Orders ({pending.length})</h3>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
          <thead><tr style={{background:'#F6F8FA'}}>
            {['Date','Party','Items','Amount',''].map((h,i)=><th key={i} style={{padding:'9px 14px',textAlign:i===3?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}
          </tr></thead>
          <tbody>
            {pending.map(o=><tr key={o.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
              <td style={{padding:'9px 14px',whiteSpace:'nowrap',color:C.sub}}>{fmtD(o.date)}</td>
              <td style={{padding:'9px 14px',fontWeight:600,color:C.txt}}>{o.partyName}</td>
              <td style={{padding:'9px 14px',fontSize:12,color:C.sub}}>
                {(o.items||[]).slice(0,2).map((x,j)=><span key={j} style={{display:'block'}}>{(x.customerName||x.itemName)} ×{x.qty}</span>)}
                {(o.items||[]).length>2&&<span style={{color:C.muted,fontSize:11}}>+{o.items.length-2} more</span>}
              </td>
              <td style={{padding:'9px 14px',textAlign:'right',fontWeight:700,color:C.blue}}>{fmt(o.total)}</td>
              <td style={{padding:'9px 14px',textAlign:'center'}}><Btn onClick={()=>openReview(o)} size="sm" variant="success">✅ Dispatched</Btn></td>
            </tr>)}
          </tbody>
        </table>
      </div>
    </div>}

    {dispatchOrder&&<DispatchPreviewModal order={dispatchOrder} onClose={()=>{setDispatchOrder(null);forceUpdate();}}/>}
    {reviewOrder&&<ReviewOrderModal order={reviewOrder} onClose={()=>setReviewOrder(null)} onSaved={sale=>{setReviewOrder(null);forceUpdate();setDoneSale(sale);}}/>}
    {doneSale&&<InvoiceModal sale={doneSale} party={state.parties.find(p=>p.id===doneSale.partyId)||{name:doneSale.partyName,phone:''}} onClose={()=>setDoneSale(null)}/>}
  </div>;
}

function EditSaleModal({sale,onClose,onSaved}){
  const[date,setDate]=useState(sale.date||today());
  const[partyId,setPartyId]=useState(sale.partyId||'');
  const[pq,setPq]=useState('');
  const[transport,setTransport]=useState(sale.transport||'');
  const[bilty,setBilty]=useState(sale.bilty||'');
  const[fMode,setFMode]=useState(sale.freightMode||'');
  const[fAmt,setFAmt]=useState(String(sale.freightAmt||''));
  const[credit,setCredit]=useState(sale.creditDays||30);
  const[lines,setLines]=useState((sale.items||[]).map((it,i)=>({id:'EL'+i,itemId:it.itemId||'',itemName:it.itemName||'',customerName:it.customerName||it.itemName||'',dispatchName:it.dispatchName,qty:it.qty||0,rate:it.rate||0,bundledJaali:it.bundledJaali})));
  const[step,setStep]=useState('item');
  const[input,setInput]=useState('');
  const[si,setSi]=useState(0);
  const[draft,setDraft]=useState({itemId:'',iName:'',cName:'',peti:1,qty:'',rate:''});
  const{toast,TC}=useToast();
  const custs=useMemo(()=>state.parties.filter(p=>p.type==='customer'),[state.parties.length]);
  const selParty=useMemo(()=>custs.find(p=>p.id===partyId),[custs,partyId]);
  const psugg=useMemo(()=>pq.length<1?[]:custs.filter(p=>p.name.toLowerCase().includes(pq.toLowerCase())).slice(0,6),[custs,pq]);
  const total=useMemo(()=>lines.reduce((s,l)=>s+(Number(l.qty)||0)*(Number(l.rate)||0),0),[lines]);
  useEffect(()=>{if(step!=='item'){setTimeout(()=>{const el=document.getElementById('edit-step-input');if(el)el.focus();},30);}else{setTimeout(()=>{const el=document.getElementById('edit-item-search');if(el)el.focus();},30);}  },[step,draft.iName]);
  const isugg=useMemo(()=>{const q=input.trim().toLowerCase();if(!q||step!=='item')return[];const words=q.split(/[\s,]+/).filter(Boolean);return state.items.filter(it=>it.active!==false).map(it=>{const cn=(it.customer||it.internal||'').toLowerCase();const hay=[cn,(it.internal||'').toLowerCase(),(it.cat||'').toLowerCase()].join(' ');let sc=0;for(const w of words){if(cn.includes(w))sc+=10;if(hay.includes(w))sc+=4;if(cn.startsWith(w))sc+=10;}if(cn.includes(q))sc+=15;return{it,sc};}).filter(x=>x.sc>0).sort((a,b)=>b.sc-a.sc).slice(0,8).map(x=>x.it);},[input,step]);
  const pickItem=it=>{const rate=selParty?getBestRate(selParty.name,it.internal||'',it.customer||'',it.sale):it.sale;setDraft({itemId:it.id,iName:it.internal||it.name||'',cName:it.customer||it.internal||'',peti:it.peti,qty:'',rate});setStep('qty');setInput('');setSi(0);};
  const commit=()=>{const qty=parseQty(draft.qty,draft.peti);if(isNaN(qty)||qty<=0)return toast('Invalid qty','error');setLines(p=>[...p,{id:uid('EL'),itemId:draft.itemId,itemName:draft.iName,customerName:draft.cName,qty,rate:Number(draft.rate)||0}]);setDraft({itemId:'',iName:'',cName:'',peti:1,qty:'',rate:''});setStep('item');setInput('');setSi(0);};
  const onKey=e=>{if(step==='item'){if(e.key==='ArrowDown'){e.preventDefault();setSi(i=>Math.min(i+1,(isugg.length||1)-1));}if(e.key==='ArrowUp'){e.preventDefault();setSi(i=>Math.max(i-1,0));}if(e.key==='Enter'&&isugg.length){e.preventDefault();pickItem(isugg[si]||isugg[0]);}if(e.key==='Escape'){setInput('');setSi(0);}}else if(step==='qty'){if(e.key==='Enter'){e.preventDefault();if(!draft.qty)return;const q=parseQty(draft.qty,draft.peti);if(isNaN(q))return toast('Invalid qty','error');setDraft(d=>({...d,qty:String(q)}));setStep('rate');}if(e.key==='Escape'){setStep('item');setInput('');}}else if(step==='rate'){if(e.key==='Enter'){e.preventDefault();commit();}if(e.key==='Escape'){setStep('qty');}}};
  const editLine=(i,k,v)=>setLines(p=>p.map((l,j)=>j!==i?l:{...l,[k]:k==='qty'||k==='rate'?Number(v)||0:v}));
  const doSave=()=>{
    if(!lines.length)return toast('Add at least one item','error');
    const updated={...sale,date,partyId,partyName:selParty?.name||sale.partyName,transport,bilty,freightMode:fMode,freightAmt:Number(fAmt)||0,creditDays:credit,items:lines.map(l=>({itemId:l.itemId,itemName:l.itemName,customerName:l.customerName,dispatchName:l.dispatchName,qty:l.qty,rate:l.rate,amount:l.qty*l.rate,bundledJaali:l.bundledJaali})),total,updatedAt:new Date().toISOString()};
    state.sales=state.sales.map(s=>s.id===sale.id?updated:s);save.sales();
    const oldQty={},newQty={};
    (sale.items||[]).forEach(l=>{oldQty[l.itemId]=(oldQty[l.itemId]||0)+l.qty;});
    lines.forEach(l=>{newQty[l.itemId]=(newQty[l.itemId]||0)+l.qty;});
    const ids=new Set([...Object.keys(oldQty),...Object.keys(newQty)]);
    ids.forEach(id=>{const delta=(newQty[id]||0)-(oldQty[id]||0); if(delta)adjustStock(id,-delta);});
    toast('Sale updated ✓');onSaved?.();onClose();
  };
  const SC={item:C.blue,qty:C.green,rate:C.amber};const SL={item:'Type item name to add…',qty:'Qty or "1 peti"',rate:'Rate ₹/pc'};
  return <Modal title={'Edit Sale — '+(sale.partyName||'?')} onClose={onClose} width={680}>
    <TC/>
    <Grid2>
      <Inp label="Date" value={date} onChange={setDate} type="date" required/>
      <Inp label="Credit Days" value={credit} onChange={v=>setCredit(Number(v)||30)} type="number"/>
      <div style={{gridColumn:'1/-1'}}>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>Party</label>
        {selParty
          ?<div style={{display:'flex',alignItems:'center',gap:10,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
            <span style={{fontWeight:700,flex:1,color:C.txt}}>{selParty.name}</span>
            <Btn onClick={()=>{setPartyId('');setPq('');}} variant="ghost" size="sm">Change</Btn>
          </div>
          :<div style={{position:'relative'}}>
            <input value={pq} autoComplete="off" onChange={e=>setPq(e.target.value)} placeholder="Type party name…"
              style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
              onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
            {psugg.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 8px 8px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
              {psugg.map(p=><div key={p.id} onMouseDown={()=>{setPartyId(p.id);setPq('');}} style={{padding:'9px 14px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`}} onMouseEnter={e=>e.currentTarget.style.background='#F9FAFB'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                <span style={{fontWeight:600,color:C.txt}}>{p.name}</span></div>)}
            </div>}
          </div>}
      </div>
      <Inp label="Transport" value={transport} onChange={setTransport} placeholder="Garage name" autoComplete="off"/>
      <Inp label="Bilty / LR" value={bilty} onChange={setBilty} placeholder="LR number" autoComplete="off"/>
    </Grid2>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',margin:'14px 0'}}>
      <div style={{padding:'10px 14px',background:SC[step]+'18',borderBottom:`1px solid ${C.bdr}`,position:'relative'}}>
        <div style={{display:'flex',gap:8,alignItems:'center'}}>
          {step!=='item'&&<span style={{fontSize:12,fontWeight:600,color:C.txt,whiteSpace:'nowrap',maxWidth:130,overflow:'hidden',textOverflow:'ellipsis'}}>{step==='rate'?`${draft.cName} × ${draft.qty}pc`:draft.cName}</span>}
          <input id={step==='item'?'edit-item-search':'edit-step-input'} key={step+draft.iName} autoComplete="off"
            value={step==='item'?input:step==='qty'?draft.qty:String(draft.rate||'')}
            onChange={e=>{const v=e.target.value;if(step==='item'){setInput(v);setSi(0);}else if(step==='qty')setDraft(d=>({...d,qty:v}));else setDraft(d=>({...d,rate:v}));}}
            onKeyDown={onKey} placeholder={SL[step]}
            style={{flex:1,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${SC[step]}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit'}}/>
          {step!=='item'&&<><Btn onClick={commit} size="sm" variant="success">Add ↵</Btn><Btn onClick={()=>{setStep('item');setInput('');}} size="sm" variant="ghost">✕</Btn></>}
        </div>
        {step==='item'&&isugg.length>0&&<div style={{position:'absolute',top:'calc(100% - 2px)',left:14,right:14,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden',maxHeight:180,display:'flex',flexDirection:'column'}}>
          <div style={{overflowY:'auto',flex:1}}>
            {isugg.map((it,i)=><div key={it.id} id={'esugg-'+i} onMouseDown={()=>{setSi(i);pickItem(it);}} onMouseEnter={()=>setSi(i)}
              style={{display:'flex',alignItems:'center',padding:'8px 14px',cursor:'pointer',background:i===si?C.blueSoft:'transparent',borderBottom:i<isugg.length-1?`1px solid ${C.bdr}`:'none'}}>
              <div style={{flex:1}}><span style={{fontWeight:600,color:C.txt,fontSize:13}}>{it.internal||it.name}</span><span style={{fontSize:11,color:C.sub,marginLeft:6}}>{it.cat}</span></div>
              <span style={{fontSize:13,fontWeight:700,color:C.blue}}>₹{selParty?getBestRate(selParty.name,it.internal||'',it.customer||'',it.sale):it.sale}/pc</span>
            </div>)}
          </div>
        </div>}
      </div>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>{['#','Item','Qty','Rate ₹','Amount',''].map((h,i)=><th key={i} style={{padding:'7px 12px',textAlign:i>1&&i<5?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}</tr></thead>
        <tbody>
          {!lines.length&&<tr><td colSpan={6} style={{padding:'20px',textAlign:'center',color:C.muted}}>No items</td></tr>}
          {lines.map((l,i)=><tr key={l.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
            <td style={{padding:'7px 12px',color:C.muted,fontSize:11}}>{i+1}</td>
            <td style={{padding:'7px 12px',fontWeight:600,color:C.txt}}>{l.customerName||l.itemName}</td>
            <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.qty} onChange={e=>editLine(i,'qty',e.target.value)} type="number" style={{width:60,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
            <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.rate} onChange={e=>editLine(i,'rate',e.target.value)} type="number" style={{width:75,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
            <td style={{padding:'7px 12px',textAlign:'right',fontWeight:700,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(l.qty*l.rate)}</td>
            <td style={{padding:'7px 12px',textAlign:'center'}}><button onClick={()=>setLines(p=>p.filter((_,j)=>j!==i))} style={{background:'none',border:'none',cursor:'pointer',color:C.muted,fontSize:15}}>×</button></td>
          </tr>)}
        </tbody>
        {lines.length>0&&<tfoot><tr style={{background:'#F6F8FA'}}><td colSpan={4} style={{padding:'9px 12px',fontWeight:700,textAlign:'right',color:C.txt}}>TOTAL</td><td style={{padding:'9px 12px',textAlign:'right',fontSize:16,fontWeight:800,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(total)}</td><td/></tr></tfoot>}
      </table>
    </div>
    <div style={{marginBottom:14}}>
      <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:6}}>Freight</label>
      <div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
        {[['topay','To Pay by party'],['','Not applicable'],['paid','Paid by us']].map(([val,lbl])=>(
          <button key={val} onClick={()=>setFMode(val)} style={{padding:'5px 13px',borderRadius:20,fontSize:12,fontWeight:600,cursor:'pointer',fontFamily:'inherit',border:`1.5px solid ${fMode===val?C.blue:C.bdr}`,background:fMode===val?C.blueSoft:'transparent',color:fMode===val?C.blue:C.sub}}>{lbl}</button>
        ))}
        {fMode==='paid'&&<input value={fAmt} onChange={e=>setFAmt(e.target.value)} type="number" placeholder="Rs amount" style={{width:90,padding:'5px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:12,outline:'none',fontFamily:'inherit'}}/>}
      </div>
    </div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <button onClick={()=>{if(window.confirm('Delete this sale?')){state.sales=state.sales.filter(s=>s.id!==sale.id);save.sales();(sale.items||[]).forEach(l=>adjustStock(l.itemId,l.qty));onSaved?.();onClose();toast('Deleted');}}} style={{padding:'7px 14px',borderRadius:7,border:`1px solid ${C.red}`,background:C.redSoft,color:C.red,cursor:'pointer',fontSize:12,fontWeight:600,fontFamily:'inherit'}}>🗑 Delete</button>
      <div style={{display:'flex',gap:8}}>
        <Btn onClick={onClose} variant="secondary">Cancel</Btn>
        <Btn onClick={doSave} disabled={!lines.length} icon="💾">Save Changes</Btn>
      </div>
    </div>
  </Modal>;
}

// ── SALES RETURN MODAL ────────────────────────────────────────────────────────

function SalesReturnModal({sale,initialParty,initialDate,onClose,onSaved}){
  // A historical (pre-app) invoice line has no itemId at all — only a name —
  // so it's resolved against today's Item Master by name here; an ERP sale
  // line already carries a real itemId and is used as-is. Either way, the
  // resolved item's category picks the sensible default return destination
  // (damage-tracked categories default to 'damaged', matching the old
  // always-damaged behavior; everything else defaults to 'stock'), which the
  // per-line selector below then lets the user override.
  const normItem=(it,i)=>{
    const isArr=Array.isArray(it);
    const name=isArr?(it[0]||''):(it.customerName||it.itemName||'');
    const qty=Number(isArr?it[1]:it.qty)||0;
    const rate=Number(isArr?it[2]:it.rate)||0;
    const itemId=(isArr?'':(it.itemId||''))||findItemByName(name)?.id||'';
    const cat=state.items.find(x=>x.id===itemId)?.cat;
    return{id:'RL'+i,itemId,itemName:name,customerName:name,qty,rate,returnQty:qty,returning:false,destination:DAMAGE_TRACKED_CATS.has(cat)?'damaged':'stock',damageNote:''};
  };
  const initLines=(sale?.items||[]).map(normItem);
  const[returnLines,setReturnLines]=useState(initLines);
  // Defaults to the date of the payment/accounts-entry this return is being
  // recorded alongside (passed in as initialDate from AccountsModule), not
  // today() — a return entered for a past collection should land on that
  // collection's date unless the user deliberately changes it. Still fully
  // editable via the field below.
  const[date,setDate]=useState(initialDate||today());
  const[reason,setReason]=useState('');
  const[party,setParty]=useState(state.parties.find(p=>p.id===sale?.partyId)||initialParty||null);
  const[pq,setPq]=useState('');
  const[selInv,setSelInv]=useState(null);
  const{toast,TC}=useToast();
  const custs=useMemo(()=>state.parties.filter(p=>p.type==='customer'),[state.parties.length]);
  const psugg=useMemo(()=>pq.length<1?[]:custs.filter(p=>p.name.toLowerCase().includes(pq.toLowerCase())).slice(0,6),[custs,pq]);
  // Merges the static pre-ERP history with sales created in-app this session
  // (state.sales) — mirroring getPartyBills()'s merge in store.js. Without
  // this, any invoice created via the in-app "New Sale" flow (e.g. one dated
  // after the static seed data ends) is structurally invisible here even
  // though it's a perfectly valid, outstanding invoice to return against.
  const partyHistory=useMemo(()=>{
    if(!party)return[];
    const names=[party.name,...(party.mergedNames||[])];
    const histBills=names.flatMap(n=>{
      const h=PARTY_HISTORY[n]||{};
      return(h.s||[]).map((s,i)=>({idx:n+'_H'+i,date:s[0],total:s[1],items:s[2]||[]}));
    });
    const newBills=state.sales.filter(s=>s.partyId===party.id&&s.type!=='return').map((s,i)=>({
      idx:'N'+i,date:s.date,total:s.total,
      items:(s.items||[]).map(x=>[x.customerName||x.itemName,x.qty,x.rate]),
    }));
    return[...histBills,...newBills].sort((a,b)=>b.date.localeCompare(a.date));
  },[party,state.sales.length]);
  const loadInvoice=inv=>{setSelInv(inv);setReturnLines(inv.items.map((x,i)=>normItem(x,i)));};
  const selected=returnLines.filter(l=>l.returning&&Number(l.returnQty)>0);
  const returnAmt=selected.reduce((s,l)=>s+(Number(l.returnQty)||0)*(Number(l.rate)||0),0);
  const doSave=()=>{
    if(!party)return toast('Select a party','error');
    if(!selected.length)return toast('Select at least one item','error');
    const ret={id:uid('RET'),type:'return',date,partyId:party.id,partyName:party.name,refSaleId:sale?.id||null,
      reason:[reason,...selected.filter(l=>l.destination==='damaged'&&l.damageNote.trim()).map(l=>(l.customerName||l.itemName)+': '+l.damageNote.trim())].filter(Boolean).join(' · '),
      items:selected.map(l=>({itemId:l.itemId,itemName:l.itemName,customerName:l.customerName||l.itemName,qty:-(Number(l.returnQty)||0),rate:Number(l.rate)||0,amount:-(Number(l.returnQty)||0)*(Number(l.rate)||0),destination:l.destination,damageNote:l.destination==='damaged'?l.damageNote.trim():undefined})),
      total:-returnAmt,createdAt:new Date().toISOString()};
    state.sales=[ret,...state.sales];save.sales();
    selected.forEach(l=>applyReturnStock(l.itemId,Number(l.returnQty)||0,l.destination));
    toast('Return recorded — '+fmt(returnAmt)+' credited');
    onSaved?.();onClose();
  };
  return <Modal title="Record Sales Return" onClose={onClose} width={560}>
    <TC/>
    {sale&&<div style={{background:C.blueSoft,borderRadius:8,padding:'10px 14px',marginBottom:16,border:`1px solid ${C.blue}20`,fontSize:12}}>
      <span style={{color:C.sub}}>Returning against: </span><span style={{fontWeight:700,color:C.txt}}>{sale.partyName}</span><span style={{color:C.sub,marginLeft:8}}>{fmtD(sale.date)} · {fmt(sale.total)}</span>
    </div>}
    <Grid2>
      <Inp label="Return Date" value={date} onChange={setDate} type="date" required/>
      <div>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>Party</label>
        {party
          ?<div style={{display:'flex',alignItems:'center',gap:8,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
            <span style={{fontWeight:700,flex:1,color:C.txt,fontSize:13}}>{party.name}</span>
            <Btn onClick={()=>{setParty(null);setPq('');setReturnLines([]);}} variant="ghost" size="sm">Change</Btn>
          </div>
          :<div style={{position:'relative'}}>
            <input value={pq} autoComplete="off" onChange={e=>setPq(e.target.value)} placeholder="Type party…"
              style={{width:'100%',padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
              onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
            {psugg.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 8px 8px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
              {psugg.map(p=><div key={p.id} onMouseDown={()=>{setParty(p);setPq('');}} style={{padding:'9px 14px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`}} onMouseEnter={e=>e.currentTarget.style.background='#F9FAFB'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}><span style={{fontWeight:600,color:C.txt}}>{p.name}</span></div>)}
            </div>}
          </div>}
      </div>
      <div style={{gridColumn:'1/-1'}}><Inp label="Reason (optional)" value={reason} onChange={setReason} placeholder="e.g. Damaged, Wrong item"/></div>
    </Grid2>
    {party&&partyHistory.length>0&&!sale&&<div style={{marginBottom:14}}>
      <div style={{fontSize:12,fontWeight:700,color:C.txt,marginBottom:8}}>Select Invoice to Return Against</div>
      <div style={{maxHeight:160,overflowY:'auto',borderRadius:8,border:`1px solid ${C.bdr}`}}>
        {partyHistory.map((inv,i)=><div key={i} onClick={()=>loadInvoice(inv)}
          style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'9px 14px',cursor:'pointer',borderBottom:i<partyHistory.length-1?`1px solid ${C.bdr}`:'none',background:selInv?.idx===inv.idx?C.blueSoft:'transparent'}}
          onMouseEnter={e=>e.currentTarget.style.background=C.blueSoft} onMouseLeave={e=>e.currentTarget.style.background=selInv?.idx===inv.idx?C.blueSoft:'transparent'}>
          <div><span style={{fontWeight:600,color:C.txt,fontSize:13}}>{fmtD(inv.date)}</span><span style={{fontSize:11,color:C.sub,marginLeft:10}}>{inv.items.slice(0,2).map(x=>Array.isArray(x)?x[0]:x).join(', ')}{inv.items.length>2?` +${inv.items.length-2} more`:''}</span></div>
          <span style={{fontWeight:700,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(inv.total)}</span>
        </div>)}
      </div>
      {selInv&&<div style={{fontSize:11,color:C.green,marginTop:6,fontWeight:600}}>✓ Items loaded from {fmtD(selInv.date)} — adjust quantities below</div>}
    </div>}
    <div style={{margin:'14px 0 4px',fontSize:12,fontWeight:700,color:C.txt}}>Items to Return</div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',marginBottom:16}}>
      {!returnLines.length&&<div style={{padding:20,textAlign:'center',color:C.muted,fontSize:13}}>{party?'Select an invoice above to load items':'Select a party first'}</div>}
      {returnLines.map((l,i)=>{
        const showDest=l.returning&&!!l.itemId;
        return <div key={l.id} style={{padding:'10px 14px',borderBottom:i<returnLines.length-1?`1px solid ${C.bdr}`:'none',opacity:l.returning?1:.6}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
        <input type="checkbox" checked={l.returning} onChange={e=>setReturnLines(p=>p.map((x,j)=>j!==i?x:{...x,returning:e.target.checked}))} style={{cursor:'pointer',width:15,height:15,flexShrink:0}}/>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontWeight:600,color:C.txt,fontSize:13,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{l.customerName||l.itemName}</div>
          <div style={{fontSize:11,color:C.sub}}>Original: {l.qty} pcs @ {fmt(l.rate)}</div>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:6,flexShrink:0}}>
          <span style={{fontSize:11,color:C.sub}}>Qty:</span>
          <input value={l.returnQty} onChange={e=>setReturnLines(p=>p.map((x,j)=>j!==i?x:{...x,returnQty:Number(e.target.value)||0}))} type="number" disabled={!l.returning} min={0} max={l.qty}
            style={{width:55,padding:'4px 6px',borderRadius:6,border:`1.5px solid ${l.returning?C.bdr:'transparent'}`,fontSize:12,textAlign:'center',outline:'none',fontFamily:'inherit'}}/>
          <span style={{fontSize:11,color:C.sub}}>@₹</span>
          <input value={l.rate} onChange={e=>setReturnLines(p=>p.map((x,j)=>j!==i?x:{...x,rate:Number(e.target.value)||0}))} type="number" disabled={!l.returning} min={0}
            title="Return rate per piece — editable if it differs from the original sale rate"
            style={{width:65,padding:'4px 6px',borderRadius:6,border:`1.5px solid ${l.returning?C.bdr:'transparent'}`,fontSize:12,textAlign:'center',outline:'none',fontFamily:'inherit'}}/>
          <span style={{fontSize:13,fontWeight:700,color:C.red,minWidth:70,textAlign:'right',fontVariantNumeric:'tabular-nums'}}>{l.returning&&Number(l.returnQty)>0?fmt(Number(l.returnQty)*Number(l.rate)):''}</span>
        </div>
        </div>
        {showDest&&<div style={{marginTop:8,paddingLeft:25}}>
          <div style={{display:'flex',alignItems:'center',gap:6}}>
            <span style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.3}}>Received as:</span>
            {[['stock','Good — Stock'],['damaged','Damaged'],['repaired','Repaired']].map(([d,label])=>
              <button key={d} onClick={()=>setReturnLines(p=>p.map((x,j)=>j!==i?x:{...x,destination:d}))}
                style={{padding:'3px 10px',borderRadius:20,border:`1.5px solid ${l.destination===d?C.blue:C.bdr}`,background:l.destination===d?C.blueSoft:'transparent',color:l.destination===d?C.blue:C.sub,fontSize:11,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>{label}</button>)}
          </div>
          {l.destination==='damaged'&&<input value={l.damageNote} onChange={e=>setReturnLines(p=>p.map((x,j)=>j!==i?x:{...x,damageNote:e.target.value}))}
            placeholder="What's damaged? e.g. Glass cracked, screen not working"
            style={{width:'100%',marginTop:6,padding:'5px 9px',borderRadius:6,border:`1.5px solid ${C.red}40`,fontSize:12,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
            onFocus={e=>e.target.style.borderColor=C.red} onBlur={e=>e.target.style.borderColor=C.red+'40'}/>}
        </div>}
      </div>;})}
    </div>
    {returnAmt>0&&<div style={{background:C.redSoft,borderRadius:8,border:`1px solid ${C.red}30`,padding:'12px 16px',marginBottom:16,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <div><div style={{fontSize:12,color:C.sub}}>Return Amount</div><div style={{fontSize:22,fontWeight:800,color:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(returnAmt)}</div></div>
      <div style={{fontSize:12,color:C.sub,textAlign:'right'}}><div>{selected.length} items · {selected.reduce((s,l)=>s+Number(l.returnQty),0)} pcs</div><div style={{marginTop:2}}>Will reduce OD by {fmt(returnAmt)}</div></div>
    </div>}
    <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
      <Btn onClick={onClose} variant="secondary">Cancel</Btn>
      <Btn onClick={doSave} disabled={!selected.length||returnAmt===0||!party} variant="danger">Record Return · −{fmt(returnAmt)}</Btn>
    </div>
  </Modal>;
}

// Edits an already-recorded return — unlike EditSaleModal (built for ordinary
// sales, which only ever touch plain stock), this needs to reverse whichever
// bucket each line originally landed in (stock/damaged/repaired) before
// re-applying the edited quantity/destination, and lets the damage remark be
// added or changed after the fact. A line with no itemId (an unresolvable
// historical row) can't be inventory-adjusted at all, so its destination
// controls are hidden, same as in SalesReturnModal.
function EditReturnModal({sale,onClose,onSaved}){
  const[date,setDate]=useState(sale.date||today());
  const[reason,setReason]=useState(sale.reason||'');
  const[lines,setLines]=useState((sale.items||[]).map((it,i)=>({
    id:'ERL'+i,
    itemId:it.itemId||'',
    itemName:it.itemName||it.customerName||'',
    customerName:it.customerName||it.itemName||'',
    qty:Math.abs(it.qty||0),
    rate:it.rate||0,
    destination:it.destination||'stock',
    damageNote:it.damageNote||'',
  })));
  const{toast,TC}=useToast();
  const total=useMemo(()=>lines.reduce((s,l)=>s+(Number(l.qty)||0)*(Number(l.rate)||0),0),[lines]);

  const doSave=()=>{
    if(!lines.length)return toast('At least one item required','error');
    // Undo the OLD inventory effect of every original line, then apply the
    // edited lines fresh — handles a qty change, a destination change, or
    // both, without needing to diff the two states line by line.
    (sale.items||[]).forEach(old=>{
      if(!old.itemId)return;
      applyReturnStock(old.itemId,-Math.abs(old.qty||0),old.destination||'stock');
    });
    lines.forEach(l=>{
      const q=Number(l.qty)||0;
      if(l.itemId&&q>0)applyReturnStock(l.itemId,q,l.destination);
    });
    const updated={...sale,date,reason,
      items:lines.map(l=>({itemId:l.itemId,itemName:l.itemName,customerName:l.customerName||l.itemName,qty:-(Number(l.qty)||0),rate:Number(l.rate)||0,amount:-(Number(l.qty)||0)*(Number(l.rate)||0),destination:l.destination,damageNote:l.destination==='damaged'?l.damageNote.trim():undefined})),
      total:-total,updatedAt:new Date().toISOString()};
    state.sales=state.sales.map(s=>s.id===sale.id?updated:s);save.sales();
    toast('Return updated ✓');onSaved?.();onClose();
  };

  const doDelete=()=>{
    if(!window.confirm('Delete this return? Inventory changes it made will be reversed.'))return;
    (sale.items||[]).forEach(old=>{
      if(!old.itemId)return;
      applyReturnStock(old.itemId,-Math.abs(old.qty||0),old.destination||'stock');
    });
    state.sales=state.sales.filter(s=>s.id!==sale.id);save.sales();
    toast('Return deleted');onSaved?.();onClose();
  };

  return <Modal title={'Edit Return — '+(sale.partyName||'?')} onClose={onClose} width={560}>
    <TC/>
    <Grid2>
      <Inp label="Return Date" value={date} onChange={setDate} type="date" required/>
      <Inp label="Reason (optional)" value={reason} onChange={setReason} placeholder="e.g. Damaged, Wrong item"/>
    </Grid2>
    <div style={{margin:'14px 0 4px',fontSize:12,fontWeight:700,color:C.txt}}>Items</div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',marginBottom:16}}>
      {!lines.length&&<div style={{padding:20,textAlign:'center',color:C.muted,fontSize:13}}>No items</div>}
      {lines.map((l,i)=>{
        const showDest=!!l.itemId;
        return <div key={l.id} style={{padding:'10px 14px',borderBottom:i<lines.length-1?`1px solid ${C.bdr}`:'none'}}>
          <div style={{display:'flex',alignItems:'center',gap:10}}>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:600,color:C.txt,fontSize:13,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{l.customerName||l.itemName}</div>
              {!l.itemId&&<div style={{fontSize:10,color:C.muted}}>Item not linked — can't adjust inventory for this line</div>}
            </div>
            <div style={{display:'flex',alignItems:'center',gap:6,flexShrink:0}}>
              <span style={{fontSize:11,color:C.sub}}>Qty:</span>
              <input value={l.qty} onChange={e=>setLines(p=>p.map((x,j)=>j!==i?x:{...x,qty:Number(e.target.value)||0}))} type="number" min={0}
                style={{width:60,padding:'4px 8px',borderRadius:6,border:`1.5px solid ${C.bdr}`,fontSize:12,textAlign:'center',outline:'none',fontFamily:'inherit'}}/>
              <span style={{fontSize:11,color:C.sub}}>@</span>
              <input value={l.rate} onChange={e=>setLines(p=>p.map((x,j)=>j!==i?x:{...x,rate:Number(e.target.value)||0}))} type="number" min={0}
                style={{width:70,padding:'4px 8px',borderRadius:6,border:`1.5px solid ${C.bdr}`,fontSize:12,textAlign:'center',outline:'none',fontFamily:'inherit'}}/>
              <span style={{fontSize:13,fontWeight:700,color:C.red,minWidth:70,textAlign:'right',fontVariantNumeric:'tabular-nums'}}>{fmt(l.qty*l.rate)}</span>
            </div>
          </div>
          {showDest&&<div style={{marginTop:8,paddingLeft:0}}>
            <div style={{display:'flex',alignItems:'center',gap:6}}>
              <span style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.3}}>Received as:</span>
              {[['stock','Good — Stock'],['damaged','Damaged'],['repaired','Repaired']].map(([d,label])=>
                <button key={d} onClick={()=>setLines(p=>p.map((x,j)=>j!==i?x:{...x,destination:d}))}
                  style={{padding:'3px 10px',borderRadius:20,border:`1.5px solid ${l.destination===d?C.blue:C.bdr}`,background:l.destination===d?C.blueSoft:'transparent',color:l.destination===d?C.blue:C.sub,fontSize:11,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>{label}</button>)}
            </div>
            {l.destination==='damaged'&&<input value={l.damageNote} onChange={e=>setLines(p=>p.map((x,j)=>j!==i?x:{...x,damageNote:e.target.value}))}
              placeholder="What's damaged? e.g. Glass cracked, screen not working"
              style={{width:'100%',marginTop:6,padding:'5px 9px',borderRadius:6,border:`1.5px solid ${C.red}40`,fontSize:12,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
              onFocus={e=>e.target.style.borderColor=C.red} onBlur={e=>e.target.style.borderColor=C.red+'40'}/>}
          </div>}
        </div>;
      })}
    </div>
    <div style={{background:C.redSoft,borderRadius:8,border:`1px solid ${C.red}30`,padding:'12px 16px',marginBottom:16,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <span style={{fontSize:12,color:C.sub}}>Return Amount</span>
      <span style={{fontSize:20,fontWeight:800,color:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(total)}</span>
    </div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <button onClick={doDelete} style={{padding:'7px 14px',borderRadius:7,border:`1px solid ${C.red}`,background:C.redSoft,color:C.red,cursor:'pointer',fontSize:12,fontWeight:600,fontFamily:'inherit'}}>🗑 Delete</button>
      <div style={{display:'flex',gap:8}}>
        <Btn onClick={onClose} variant="secondary">Cancel</Btn>
        <Btn onClick={doSave} disabled={!lines.length}>Save Changes</Btn>
      </div>
    </div>
  </Modal>;
}


function PartyProfile({party,onClose}){
  const mobile=useIsMobile();
  const hist=PARTY_HISTORY[party.name]||{};
  const sales=hist.s||[];
  const pmts=hist.p||[];
  const st=gPS(party.name);
  const newSales=state.sales.filter(s=>s.partyId===party.id).reduce((s,t)=>s+t.total,0);
  const totalSold=(st?.ts||0)+newSales;
  const currentOD=getCustomerBal(party.id);
  const odRatio=totalSold>0?currentOD/totalSold:0;
  const grade=currentOD<=0?'A':odRatio<0.15?'A':odRatio<0.3?'B':odRatio<0.6?'C':'D';
  const gradeColor={A:C.green,B:C.blue,C:C.amber,D:C.red};
  const gradeLabel={A:'Excellent payer',B:'Good — minor delays',C:'Average — needs follow-up',D:'At Risk — high overdue'};
  const firstDate=sales.length?sales[sales.length-1][0]:null;
  const monthsActive=firstDate?Math.round((new Date('2026-07-11')-new Date(firstDate+'T00:00:00'))/(30*86400000)):0;
  const avgDaysBetweenPayments=pmts.length>1?pmts.slice(0,-1).reduce((s,p,i)=>s+Math.abs((new Date(p[0])-new Date(pmts[i+1][0]))/86400000),0)/Math.max(1,pmts.length-1):0;
  const[tab,setTab]=useState('overview');
  const[aiText,setAiText]=useState('');
  const[aiLoading,setAiLoading]=useState(false);
  const generateAI=async()=>{
    setAiLoading(true);setAiText('');
    const prompt=`You are a business analyst for Upkaran, a kitchenware distributorship in Raipur, CG.\n\nAnalyze this customer and write a concise business intelligence summary (4-5 sentences):\n\nCustomer: ${party.name}\nCity: ${party.city||'—'} | Route: ${party.route||'—'}\nActive since: ${firstDate||'Unknown'} (${monthsActive} months)\nTotal Business: Rs${totalSold.toLocaleString('en-IN')}\nAvg Monthly: Rs${(st?.am||0).toLocaleString('en-IN')}\nCurrent OD: Rs${currentOD.toLocaleString('en-IN')}\nGrade: ${grade} — ${gradeLabel[grade]}\nRecent sales: ${sales.slice(0,3).map(s=>s[0]+' Rs'+s[1].toLocaleString('en-IN')).join(' | ')}\nRecent payments: ${pmts.slice(0,3).map(p=>p[0]+' Rs'+p[1].toLocaleString('en-IN')+' ('+p[2]+')').join(' | ')}\n\nWrite a practical 4-5 sentence summary: business value, payment behaviour, relationship health, and one actionable recommendation. Be specific with numbers. English.`;
    try{
      const res=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({model:'claude-sonnet-4-6',max_tokens:400,messages:[{role:'user',content:prompt}]})});
      const data=await res.json();
      setAiText(data.content?.[0]?.text||'Unable to generate summary.');
    }catch(e){setAiText('Error: '+e.message);}
    setAiLoading(false);
  };
  const TABS=[['overview','Overview'],['sales','Sales History'],['payments','Payments'],['ai','AI Summary']];
  return <Modal title="" onClose={onClose} width={640}>
    <div style={{background:'linear-gradient(135deg,#0A0E1A,#1a2035)',borderRadius:10,padding:'20px 22px',marginBottom:20}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start'}}>
        <div>
          <div style={{fontSize:18,fontWeight:800,color:'#fff',marginBottom:3}}>{party.name}</div>
          <div style={{fontSize:12,color:'rgba(255,255,255,.5)'}}>{party.city} · {party.route} · {party.credit}d credit</div>
          {party.phone&&<div style={{fontSize:12,color:'rgba(255,255,255,.5)',marginTop:2}}>📞 {party.phone}</div>}
        </div>
        <div style={{textAlign:'right'}}>
          <div style={{fontSize:28,fontWeight:900,color:gradeColor[grade],lineHeight:1}}>{grade}</div>
          <div style={{fontSize:10,color:'rgba(255,255,255,.4)',marginTop:2}}>{gradeLabel[grade]}</div>
        </div>
      </div>
      <div style={{display:'grid',gridTemplateColumns:mobile?'repeat(2,1fr)':'repeat(4,1fr)',gap:10,marginTop:16}}>
        {[['Total',fmtL(totalSold),C.blue],['Avg/Mo',st?.am?fmtL(st.am):'—',C.purple],['OD',currentOD<0?'−'+fmtL(-currentOD):fmtL(currentOD),currentOD>0?C.red:C.green],['Active',monthsActive+'mo',C.amber]].map(([l,v,col],i)=>(
          <div key={i} style={{background:'rgba(255,255,255,.07)',borderRadius:8,padding:'9px 12px'}}>
            <div style={{fontSize:9,color:'rgba(255,255,255,.4)',textTransform:'uppercase',letterSpacing:.6,marginBottom:3}}>{l}</div>
            <div style={{fontSize:16,fontWeight:800,color:col,fontVariantNumeric:'tabular-nums'}}>{v}</div>
          </div>))}
      </div>
    </div>
    <div style={{display:'flex',borderBottom:`1px solid ${C.bdr}`,marginBottom:18,gap:mobile?0:2,overflowX:mobile?'auto':'visible'}}>
      {TABS.map(([id,label])=><button key={id} onClick={()=>setTab(id)} style={{padding:mobile?'8px 10px':'8px 16px',border:'none',background:'transparent',cursor:'pointer',fontFamily:'inherit',fontSize:mobile?12:13,whiteSpace:'nowrap',fontWeight:tab===id?700:400,color:tab===id?C.blue:C.sub,borderBottom:tab===id?`2px solid ${C.blue}`:'2px solid transparent',marginBottom:-1}}>{label}</button>)}
    </div>
    {tab==='overview'&&<div>
      <div style={{display:'grid',gridTemplateColumns:mobile?'1fr':'1fr 1fr',gap:14,marginBottom:18}}>
        <div style={{background:'#F9FAFB',borderRadius:8,padding:14,border:`1px solid ${C.bdr}`}}>
          <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:10}}>Business Relationship</div>
          {[['Customer since',firstDate?fmtD(firstDate):'—'],['Active for',monthsActive>0?monthsActive+' months':'—'],['Total invoices',sales.length+' recorded'],['Total payments',pmts.length+' recorded'],['Avg payment gap',avgDaysBetweenPayments>0?Math.round(avgDaysBetweenPayments)+'d':'—']].map(([l,v])=>(
            <div key={l} style={{display:'flex',justifyContent:'space-between',padding:'5px 0',borderBottom:`1px solid ${C.bdr}`,fontSize:12}}>
              <span style={{color:C.sub}}>{l}</span><span style={{fontWeight:600,color:C.txt}}>{v}</span>
            </div>))}
        </div>
        <div style={{background:'#F9FAFB',borderRadius:8,padding:14,border:`1px solid ${C.bdr}`}}>
          <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:10}}>Items Bought</div>
          {[...new Set(sales.flatMap(s=>(s[2]||[]).map(x=>Array.isArray(x)?x[0]:x)))].slice(0,8).map(item=>(
            <div key={item} style={{padding:'4px 0',fontSize:12,color:C.txt,borderBottom:`1px solid ${C.bdr}`}}>{item}</div>))}
        </div>
      </div>
      {(()=>{
        const bills=getPartyBills(party);
        if(!bills.length) return null;
        return <div>
          <div style={{fontSize:12,fontWeight:700,color:C.txt,marginBottom:8}}>Outstanding Invoices
            <span style={{fontSize:11,fontWeight:400,color:C.sub,marginLeft:8}}>FIFO — oldest bills settled first</span>
          </div>
          {bills.map(b=>{
            const paid=b.paid;
            return <div key={b.key} style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'8px 12px',borderRadius:7,marginBottom:6,
              border:`1.5px solid ${paid?C.green:b.overdue>0?C.red+'40':C.bdr}`,background:paid?C.greenSoft:b.overdue>0?C.redSoft:'#F9FAFB'}}>
              <div><span style={{fontWeight:600,color:C.txt,fontSize:13}}>{fmtD(b.date)}</span>
                <span style={{fontSize:11,fontWeight:700,marginLeft:10,padding:'2px 8px',borderRadius:20,background:paid?C.green:b.overdue>0?C.red:C.green,color:'#fff'}}>
                  {paid?'Paid ✓':b.overdue>0?b.overdue+'d overdue':'In credit'}
                </span>
                {b.isNew&&<span style={{fontSize:10,fontWeight:700,marginLeft:6,color:C.blue}}>NEW</span>}
                <div style={{fontSize:11,color:C.sub,marginTop:2}}>{b.items||'—'}</div>
              </div>
              <div style={{textAlign:'right'}}>
                {!paid&&b.remaining!==b.orig&&<div style={{fontSize:10,color:C.sub,textDecoration:'line-through'}}>{fmt(b.orig)}</div>}
                <div style={{fontWeight:800,color:paid?C.green:C.red}}>{paid?'₹0':fmt(b.remaining)}</div>
                <div style={{fontSize:11,color:C.muted}}>billed {fmt(b.total)}</div>
              </div>
            </div>;
          })}
        </div>;
      })()}
    </div>}
    {tab==='sales'&&<div style={{maxHeight:380,overflowY:'auto'}}>
      <CardTable mobile={mobile} emptyLabel="No sales history"
        rows={[...state.sales.filter(s=>s.partyId===party.id).map(s=>({d:s.date,t:s.total,items:s.items?.map(x=>([x.customerName||x.itemName,x.qty,x.rate]))||[],isNew:true})),...sales.map(s=>({d:s[0],t:s[1],items:s[2]||[],isNew:false}))].sort((a,b)=>b.d.localeCompare(a.d))}
        keyFn={(s,i)=>i} rowStyle={s=>s.isNew?{background:C.blueSoft}:null}
        footer={<div style={{display:'flex',justifyContent:'space-between',padding:'9px 12px',background:'#F6F8FA',fontWeight:700,color:C.txt}}>
          <span>Total</span><span style={{fontWeight:800,color:C.blue}}>{fmt(totalSold)}</span>
        </div>}
        columns={[
          {key:'date',label:'Date',render:s=><span style={{whiteSpace:'nowrap'}}>{fmtD(s.d)}{s.isNew&&<span style={{fontSize:10,color:C.blue,fontWeight:700,marginLeft:6}}>NEW</span>}</span>},
          {key:'items',label:'Items',render:s=><span style={{fontSize:12,color:C.sub}}>{s.items.length?s.items.slice(0,3).map((x,j)=><span key={j} style={{display:'block'}}>{Array.isArray(x)?x[0]+' ×'+x[1]+' @₹'+x[2]:(x.customerName||x.itemName)}</span>):'—'}</span>},
          {key:'amt',label:'Amount',align:'right',render:s=><span style={{fontWeight:700,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(s.t)}</span>},
        ]}/>
    </div>}
    {tab==='payments'&&<div style={{maxHeight:380,overflowY:'auto'}}>
      <CardTable mobile={mobile} emptyLabel="No payment history"
        rows={[...state.payments.filter(p=>p.partyId===party.id).map(p=>({d:p.date,a:p.amount,m:'ERP',isNew:true})),...pmts.map(p=>({d:p[0],a:p[1],m:p[2],isNew:false}))].sort((a,b)=>b.d.localeCompare(a.d))}
        keyFn={(p,i)=>i} rowStyle={p=>p.isNew?{background:C.greenSoft}:null}
        footer={<div style={{display:'flex',justifyContent:'space-between',padding:'9px 12px',background:'#F6F8FA',fontWeight:700,color:C.txt}}>
          <span>Total Received</span><span style={{fontWeight:800,color:C.green}}>{fmt(pmts.reduce((s,p)=>s+p[1],0)+state.payments.filter(p=>p.partyId===party.id).reduce((s,p)=>s+p.amount,0))}</span>
        </div>}
        columns={[
          {key:'date',label:'Date',render:p=><span style={{whiteSpace:'nowrap'}}>{fmtD(p.d)}{p.isNew&&<span style={{fontSize:10,color:C.green,fontWeight:700,marginLeft:6}}>NEW</span>}</span>},
          {key:'mode',label:'Mode',render:p=><Badge v={p.m} color={p.m==='Cash'?'amber':p.m==='Bank'?'blue':'green'}/>},
          {key:'amt',label:'Amount',align:'right',render:p=><span style={{fontWeight:700,color:C.green,fontVariantNumeric:'tabular-nums'}}>{fmt(p.a)}</span>},
        ]}/>
    </div>}
    {tab==='ai'&&<div>
      {!aiText&&!aiLoading&&<div style={{textAlign:'center',padding:'30px 20px'}}>
        <div style={{fontSize:32,marginBottom:12}}>🤖</div>
        <div style={{fontSize:14,fontWeight:600,color:C.txt,marginBottom:6}}>AI Business Intelligence</div>
        <div style={{fontSize:13,color:C.sub,marginBottom:20}}>Smart analysis of payment behavior, business relationship, and actionable recommendations.</div>
        <Btn onClick={generateAI} size="lg">Generate AI Summary</Btn>
      </div>}
      {aiLoading&&<div style={{textAlign:'center',padding:'40px'}}>
        <div style={{width:32,height:32,borderRadius:'50%',border:`3px solid ${C.bdr}`,borderTopColor:C.blue,animation:'spin 0.8s linear infinite',margin:'0 auto 12px'}}/>
        <div style={{fontSize:13,color:C.sub}}>Analysing customer data…</div>
        <style>{"@keyframes spin{to{transform:rotate(360deg)}}"}</style>
      </div>}
      {aiText&&<div>
        <div style={{background:'linear-gradient(135deg,#EFF6FF,#F5F3FF)',borderRadius:10,padding:18,border:`1px solid ${C.blue}20`,marginBottom:14,fontSize:13,lineHeight:1.8,color:C.txt}}>{aiText}</div>
        <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
          <Btn onClick={()=>{setAiText('');setAiLoading(false);}} variant="secondary" size="sm">Regenerate</Btn>
          <button onClick={()=>navigator.clipboard?.writeText(aiText)} style={{padding:'6px 12px',borderRadius:6,border:`1px solid ${C.bdr}`,background:C.surf,fontSize:12,cursor:'pointer',fontFamily:'inherit',color:C.sub}}>📋 Copy</button>
        </div>
      </div>}
    </div>}
  </Modal>;
}

function PartyLedger({party,onClose}){
  const mobile=useIsMobile();
  const isSupplier=party.type==='supplier';
  const isTransporterParty=party.type==='transporter';
  const PAGE=20;
  const[actPage,setActPage]=useState(0);
  const[histPage,setHistPage]=useState(0);
  const[viewSale,setViewSale]=useState(null);
  const[viewPurchase,setViewPurchase]=useState(null);
  const[showRepairModal,setShowRepairModal]=useState(false);
  const[exFrom,setExFrom]=useState('');
  const[exTo,setExTo]=useState(today());
  const{toast,TC}=useToast();
  const repairExps=useMemo(()=>state.repairExpenses.filter(r=>r.partyId===party.id),[party.id,state.repairExpenses.length]);
  const repairExpTotal=repairExps.reduce((s,r)=>s+(r.expenseAmount||0),0);

  const pmtDesc = p => p.notes || ((isSupplier||isTransporterParty?'Paid — ':'Received — ')+(ACC_MAP[p.account]?.short||p.account||'—'));

  const newSales=isSupplier
    ?state.purchases.filter(s=>s.partyId===party.id).map(s=>({
        date:s.date, type:s.total<0?'return':'purchase', amount:s.total,
        desc:(s.items||[]).slice(0,3).map(x=>(x.supplierName||x.customerName||x.itemName)+' ×'+x.qty).join(', ')||(s.total<0?'Purchase Return':'Purchase'),
        createdAt:s.createdAt||'', purchaseId:s.id,
      }))
    :isTransporterParty
    ?state.freightCharges.filter(f=>f.transporterId===party.id).map(f=>({
        date:f.date, type:'freight', amount:f.amount,
        desc:f.notes||(f.totalPetis.toFixed(2)+' petis × ₹'+f.ratePerPeti+'/peti'),
        createdAt:f.createdAt||'',
      }))
    :state.sales.filter(s=>s.partyId===party.id).map(s=>({
        date:s.date, type:s.type==='return'?'return':'sale', amount:s.total,
        desc:(s.items||[]).slice(0,3).map(x=>(x.customerName||x.itemName)+' ×'+x.qty).join(', ')||(s.type==='return'?'Sales Return':'Sale'),
        createdAt:s.createdAt||'', saleId:s.id,
      }));
  const newPmts=((isSupplier||isTransporterParty)?state.supplierPayments:state.payments).filter(p=>p.partyId===party.id).map(p=>({
    date:p.date, type:'payment', amount:p.amount, desc:pmtDesc(p), createdAt:'',
  }));

  // Historical (pre-ERP) source: customers/suppliers only — no historical freight data exists.
  const histSalesRaw = isSupplier ? (SUPPLIER_HISTORY[party.name]?.s||[])
    : isTransporterParty ? []
    : (PARTY_HISTORY[party.name]?.s||[]);
  const histPmtsRaw = isSupplier ? (SUPPLIER_HISTORY[party.name]?.p||[])
    : isTransporterParty ? []
    : (PARTY_HISTORY[party.name]?.p||[]);
  const histSales=isSupplier
    ?histSalesRaw.map(s=>({date:s[0],type:'purchase',amount:s[1],
        desc:(s[2]||[]).slice(0,3).map(x=>Array.isArray(x)?x[0]+' ×'+x[1]:x).join(', ')||'Purchase'}))
    :histSalesRaw.flatMap(s=>splitPosNegItems(s[1],s[2]).map(seg=>({
        date:s[0],type:seg.isReturn?'return':'sale',amount:seg.total,
        desc:seg.items.slice(0,3).map(x=>Array.isArray(x)?x[0]+' ×'+x[1]:x).join(', ')||(seg.isReturn?'Sales Return':'Sale'),
        items:seg.items,
      })));
  const histPmts=histPmtsRaw.map(p=>({date:p[0],type:'payment',amount:p[1],
    desc:(isSupplier?'Paid — ':'Received — ')+(p[2]||'Cash/Bank')}));

  // Single continuous running balance across BOTH historical and in-app
  // entries, sorted purely by date, rather than two independently-scaled
  // walks (history from 0, in-app from party.ob). Walking history alone from
  // 0 always lands exactly on party.ob (verified: that's the definition of
  // an opening balance), so merging everything into one walk from 0 (or from
  // party.ob when a party has no history at all) reproduces the correct
  // balance at every point in time — including for a party like one whose
  // "historical" data keeps getting appended with recent transactions well
  // past when in-app entries also started. With two separate phases, an
  // in-app entry backdated into a date range where history is still being
  // added landed on a totally different numeric scale than its chronological
  // neighbours in the History table, so it looked like it had no effect on
  // anything around it even though the final total was already correct.
  const hasHist=histSales.length>0||histPmts.length>0;
  const histTagged=[...histSales,...histPmts].map(e=>({...e,_src:'hist'}));
  const appTagged=[...newSales,...newPmts].map(e=>({...e,_src:'app'}));
  const allTagged=[...histTagged,...appTagged].sort((a,b)=>a.date.localeCompare(b.date)||(a.createdAt||'').localeCompare(b.createdAt||''));
  // History doesn't always net to exactly party.ob — a missing/mis-keyed
  // pre-ERP entry can leave a gap (e.g. Mohan Electricals' raw history nets
  // ₹6,000 higher than its party.ob). Reconcile that gap into the STARTING
  // balance rather than assuming a perfect 0 -> party.ob walk, mirroring the
  // obExtra pattern already used in getPartyBills() for the same reason.
  // Without this, the closing balance still happens to land on party.ob
  // (app entries are anchored to it independently elsewhere), but every
  // *historical* row's own running balance was off by the gap — the History
  // table showed ₹63,800 while the party list correctly showed ₹57,800.
  const histNet=histTagged.reduce((s,e)=>s+(e.type==='payment'?-e.amount:e.amount),0);
  const obGap=hasHist?(party.ob||0)-histNet:0;
  let bal=hasHist?obGap:(party.ob||0);
  const allWithBal=allTagged.map(e=>{
    const delta=e.type==='payment'?-e.amount:e.amount;
    bal+=delta;
    return {...e,delta,balance:bal};
  });
  const currentOD=bal;
  const rows=allWithBal.filter(e=>e._src==='app').slice().reverse();
  const histRows=allWithBal.filter(e=>e._src==='hist').slice().reverse();

  const totalSold=histSales.reduce((s,e)=>s+e.amount,0)+newSales.filter(e=>e.type==='sale'||e.type==='purchase'||e.type==='freight').reduce((s,e)=>s+e.amount,0);
  const totalReturned=Math.abs(newSales.filter(e=>e.type==='return').reduce((s,e)=>s+e.amount,0))+Math.abs(histSales.filter(e=>e.type==='return').reduce((s,e)=>s+e.amount,0));
  const totalReceived=histPmts.reduce((s,e)=>s+e.amount,0)+newPmts.reduce((s,e)=>s+e.amount,0);

  const metrics=useMemo(()=>{
    if(isSupplier) return getAllVendorMetrics().find(m=>m.party.id===party.id);
    if(isTransporterParty) return null;
    return getAllCustomerMetrics().find(m=>m.party.id===party.id);
  },[party.id,isSupplier,isTransporterParty,state.sales.length,state.payments.length,state.purchases.length,state.supplierPayments.length]);

  const typeMeta={sale:{label:'Sale',color:C.blue,badge:'blue'},purchase:{label:'Purchase',color:C.blue,badge:'blue'},freight:{label:'Freight',color:C.blue,badge:'blue'},return:{label:'Return',color:C.amber,badge:'amber'},payment:{label:'Payment',color:C.green,badge:'green'}};
  const typeLabel={sale:'Sale',purchase:'Purchase',return:'Return',freight:'Freight'};
  const gradeColor={A:C.green,B:C.blue,C:C.amber,D:C.red};

  // Export source — same in-app + historical transactions as newSales/histSales
  // above, but keeping full per-line item detail (qty/rate/amount) instead of
  // the 3-item truncated `desc` string built for the on-screen tables, since a
  // downloaded ledger needs every line item, not a preview.
  const exportEntries=useMemo(()=>{
    const lineOf=x=>Array.isArray(x)?{name:x[0],qty:x[1],rate:x[2],amount:(x[1]!=null&&x[2]!=null)?x[1]*x[2]:null}:{name:String(x),qty:null,rate:null,amount:null};
    const inAppTx=isSupplier
      ?state.purchases.filter(s=>s.partyId===party.id).map(s=>({date:s.date,type:s.total<0?'return':'purchase',amount:s.total,
          items:(s.items||[]).map(x=>({name:x.supplierName||x.customerName||x.itemName||'Item',qty:x.qty,rate:x.rate,amount:x.amount??(x.qty*x.rate)}))}))
      :isTransporterParty
      ?state.freightCharges.filter(f=>f.transporterId===party.id).map(f=>({date:f.date,type:'freight',amount:f.amount,
          items:[{name:f.notes||'Freight',qty:f.totalPetis,rate:f.ratePerPeti,amount:f.amount}]}))
      :state.sales.filter(s=>s.partyId===party.id).map(s=>({date:s.date,type:s.type==='return'?'return':'sale',amount:s.total,
          items:(s.items||[]).map(x=>({name:x.customerName||x.itemName||'Item',qty:x.qty,rate:x.rate,amount:x.amount??(x.qty*x.rate)}))}));
    const inAppPmts=((isSupplier||isTransporterParty)?state.supplierPayments:state.payments).filter(p=>p.partyId===party.id)
      .map(p=>({date:p.date,type:'payment',amount:p.amount,desc:pmtDesc(p)}));

    const histTxRaw=isSupplier?(SUPPLIER_HISTORY[party.name]?.s||[]):isTransporterParty?[]:(PARTY_HISTORY[party.name]?.s||[]);
    const histTx=isSupplier
      ?histTxRaw.map(s=>({date:s[0],type:s[1]<0?'return':'purchase',amount:s[1],items:(s[2]||[]).map(lineOf)}))
      :histTxRaw.flatMap(s=>splitPosNegItems(s[1],s[2]).map(seg=>({date:s[0],type:seg.isReturn?'return':'sale',amount:seg.total,items:seg.items.map(lineOf)})));
    const histPmtsRaw2=isSupplier?(SUPPLIER_HISTORY[party.name]?.p||[]):isTransporterParty?[]:(PARTY_HISTORY[party.name]?.p||[]);
    const histPmts2=histPmtsRaw2.map(p=>({date:p[0],type:'payment',amount:p[1],desc:(isSupplier?'Paid — ':'Received — ')+(p[2]||'Cash/Bank')}));

    // ONE continuous running-balance walk across history + in-app together,
    // sorted purely by date — mirrors the same fix applied to the on-screen
    // Recent Activity/History tables above. Walking history alone from 0
    // always lands exactly on party.ob (that's the definition of an opening
    // balance — verified for every party), so merging everything into a
    // single walk from 0 (or from party.ob when there's no history at all)
    // reproduces the correct balance at every point without double-counting.
    // An earlier version of this kept two separately-scaled walks (history
    // from 0, in-app from party.ob) then merely sorted the two lists together
    // for display — correct for the closing total, but any in-app entry
    // backdated into a date range where history was still being appended
    // (e.g. a supplier whose historical file keeps getting new entries well
    // past when in-app entries also started) landed on a different numeric
    // scale than its chronological neighbours, so it appeared to have no
    // effect on the running balance even though the final total was right.
    const hasHist=histTx.length>0||histPmts2.length>0;
    const combined=[...histTx,...histPmts2,...inAppTx,...inAppPmts].sort((a,b)=>a.date.localeCompare(b.date));
    // Same reconciliation as the on-screen walk above — history doesn't
    // always net to exactly party.ob, so the gap is folded into the starting
    // balance instead of assumed away. See getPartyBills()'s obExtra for the
    // same pattern.
    const histNet=[...histTx,...histPmts2].reduce((s,e)=>s+(e.type==='payment'?-e.amount:e.amount),0);
    const obGap=hasHist?(party.ob||0)-histNet:0;
    let bal=hasHist?obGap:(party.ob||0);
    const all=combined.map(e=>{
      const delta=e.type==='payment'?-e.amount:e.amount;
      bal+=delta;
      return {...e,delta,balance:bal};
    });

    return {all,hasHist,obGap};
  },[party.id,party.name,isSupplier,isTransporterParty,state.sales.length,state.purchases.length,state.payments.length,state.supplierPayments.length,state.freightCharges.length]);

  // Slices the pre-walked [all] list (see exportEntries above) to the picked
  // [exFrom,exTo] window. "Opening balance" for the window is whatever the
  // last entry strictly before exFrom landed on — or, if the window starts
  // before any recorded activity, 0 when there's history to walk from (the
  // historical walk's own starting point) or party.ob when there's no
  // history at all (an explicit starting balance entered for the party).
  const buildExportWindow=()=>{
    const {all,hasHist,obGap}=exportEntries;
    const trueOpening=hasHist?obGap:(party.ob||0);
    const before=exFrom?all.filter(e=>e.date<exFrom):[];
    const openingBalance=before.length?before[before.length-1].balance:trueOpening;
    const filtered=all.filter(e=>(!exFrom||e.date>=exFrom)&&(!exTo||e.date<=exTo));
    return {filtered,openingBalance,closingBalance:filtered.length?filtered[filtered.length-1].balance:openingBalance};
  };

  const exportLedgerExcel=()=>{
    const {filtered,openingBalance,closingBalance}=buildExportWindow();
    if(!filtered.length)return toast('No transactions in this date range','error');
    const aoa=[
      ['Upkaran ERP — Party Ledger'],
      [party.name],
      ['Period: '+(exFrom?fmtD(exFrom):'Beginning')+' to '+(exTo?fmtD(exTo):'Today')],
      [],
      ['Date','Type','Particulars','Qty','Rate','Amount','Balance'],
      ['','','Opening Balance','','','',openingBalance],
    ];
    for(const e of filtered){
      if(e.type==='payment'){
        aoa.push([fmtD(e.date),'Payment',e.desc||'Payment','','',e.delta,e.balance]);
        continue;
      }
      const label=typeLabel[e.type]||'Sale';
      const items=e.items||[];
      if(items.length<=1){
        const it=items[0];
        aoa.push([fmtD(e.date),label,it?it.name:label,it?.qty??'',it?.rate??'',e.amount,e.balance]);
      }else{
        items.forEach((it,i)=>aoa.push([i===0?fmtD(e.date):'',i===0?label:'',it.name,it.qty??'',it.rate??'',it.amount??'','']));
        aoa.push(['','','Invoice Total','','',e.amount,e.balance]);
      }
    }
    aoa.push([]);
    aoa.push(['','','Closing Balance','','','',closingBalance]);

    const ws=XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols']=[{wch:12},{wch:10},{wch:34},{wch:8},{wch:10},{wch:12},{wch:12}];
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:6}},{s:{r:1,c:0},e:{r:1,c:6}},{s:{r:2,c:0},e:{r:2,c:6}}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Ledger');
    XLSX.writeFile(wb,party.name.replace(/[^a-z0-9]+/gi,'-')+'-ledger-'+today()+'.xlsx');
    toast('Ledger exported ✓');
  };

  const exportLedgerPDF=()=>{
    const {filtered,openingBalance,closingBalance}=buildExportWindow();
    if(!filtered.length)return toast('No transactions in this date range','error');
    const rs=n=>n===''?'':'Rs '+Number(n||0).toLocaleString('en-IN');
    const body=[['','','Opening Balance','','','',rs(openingBalance)]];
    for(const e of filtered){
      if(e.type==='payment'){
        body.push([fmtD(e.date),'Payment',e.desc||'Payment','','',rs(e.delta),rs(e.balance)]);
        continue;
      }
      const label=typeLabel[e.type]||'Sale';
      const items=e.items||[];
      if(items.length<=1){
        const it=items[0];
        body.push([fmtD(e.date),label,it?it.name:label,it?.qty??'',it?.rate!=null?rs(it.rate):'',rs(e.amount),rs(e.balance)]);
      }else{
        items.forEach((it,i)=>body.push([i===0?fmtD(e.date):'',i===0?label:'',it.name,it.qty??'',it.rate!=null?rs(it.rate):'',it.amount!=null?rs(it.amount):'','']));
        body.push(['','','Invoice Total','','',rs(e.amount),rs(e.balance)]);
      }
    }
    body.push(['','','Closing Balance','','','',rs(closingBalance)]);

    const doc=new jsPDF({orientation:'landscape',unit:'pt',format:'a4'});
    doc.setFontSize(16);doc.setFont(undefined,'bold');
    doc.text('Upkaran ERP - Party Ledger',40,40);
    doc.setFontSize(11);doc.setFont(undefined,'normal');
    doc.text(party.name,40,58);
    doc.setFontSize(9);
    doc.text('Period: '+(exFrom?fmtD(exFrom):'Beginning')+' to '+(exTo?fmtD(exTo):'Today'),40,72);
    autoTable(doc,{
      startY:84,
      head:[['Date','Type','Particulars','Qty','Rate','Amount','Balance']],
      body,
      styles:{fontSize:8,cellPadding:4},
      headStyles:{fillColor:[37,99,235],textColor:255,fontStyle:'bold'},
      columnStyles:{3:{halign:'right'},4:{halign:'right'},5:{halign:'right'},6:{halign:'right'}},
      margin:{left:40,right:40},
      didParseCell:(data)=>{
        const label=data.row.raw[2];
        if(label==='Opening Balance'||label==='Closing Balance'||label==='Invoice Total'){
          data.cell.styles.fontStyle='bold';
          data.cell.styles.fillColor=[243,244,246];
        }
      },
    });
    doc.save(party.name.replace(/[^a-z0-9]+/gi,'-')+'-ledger-'+today()+'.pdf');
    toast('Ledger exported ✓');
  };

  const actPageRows=rows.slice(actPage*PAGE,(actPage+1)*PAGE);
  const actTotalPages=Math.ceil(rows.length/PAGE);
  const histPageRows=histRows.slice(histPage*PAGE,(histPage+1)*PAGE);
  const histTotalPages=Math.ceil(histRows.length/PAGE);

  return <Modal title={party.name+' — Ledger'} onClose={onClose} width={680}>
    <TC/>
    {viewSale&&<InvoiceModal sale={viewSale} party={party} onClose={()=>setViewSale(null)} viewOnly/>}
    {viewPurchase&&<PurchaseViewModal purchase={viewPurchase} onClose={()=>setViewPurchase(null)}/>}
    {showRepairModal&&<RepairExpenseModal party={party} onClose={()=>setShowRepairModal(false)} onSaved={()=>setShowRepairModal(false)}/>}
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
      <div>
        <div style={{fontSize:11,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>{currentOD<0?'In Credit':(isSupplier||isTransporterParty)?'Current Payable':'Current Outstanding'}</div>
        <div style={{fontSize:28,fontWeight:800,color:currentOD>0?C.red:C.green,fontVariantNumeric:'tabular-nums'}}>{currentOD<0?'−':''}{fmt(Math.abs(currentOD))}</div>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,textAlign:'right'}}>
        <div><div style={{fontSize:10,color:C.sub}}>{isSupplier?'Total Purchased':isTransporterParty?'Total Freight':'Total Sold'}</div><div style={{fontWeight:700,color:C.txt}}>{fmt(totalSold)}</div></div>
        <div><div style={{fontSize:10,color:C.sub}}>{(isSupplier||isTransporterParty)?'Total Paid':'Total Received'}</div><div style={{fontWeight:700,color:C.green}}>{fmt(totalReceived)}</div></div>
        {totalReturned>0&&<div style={{gridColumn:'1/-1'}}><div style={{fontSize:10,color:C.sub}}>Total Returned</div><div style={{fontWeight:700,color:C.amber}}>{fmt(totalReturned)}</div></div>}
      </div>
    </div>

    <div style={{display:'flex',alignItems:mobile?'stretch':'flex-end',gap:8,flexWrap:'wrap',flexDirection:mobile?'column':'row',marginBottom:16,background:'#F9FAFB',border:`1px solid ${C.bdr}`,borderRadius:8,padding:'10px 12px'}}>
      <div style={{display:'flex',gap:8,width:mobile?'100%':'auto',flex:mobile?'none':undefined}}>
        <div style={{flex:1,minWidth:110}}>
          <label style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>From</label>
          <input type="date" value={exFrom} onChange={e=>setExFrom(e.target.value)} max={exTo||undefined}
            style={{width:'100%',padding:'6px 8px',borderRadius:6,border:`1.5px solid ${C.bdr}`,fontSize:12,fontFamily:'inherit',boxSizing:'border-box'}}/>
        </div>
        <div style={{flex:1,minWidth:110}}>
          <label style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>To</label>
          <input type="date" value={exTo} onChange={e=>setExTo(e.target.value)} min={exFrom||undefined}
            style={{width:'100%',padding:'6px 8px',borderRadius:6,border:`1.5px solid ${C.bdr}`,fontSize:12,fontFamily:'inherit',boxSizing:'border-box'}}/>
        </div>
      </div>
      <div style={{display:'flex',gap:8,width:mobile?'100%':'auto'}}>
        <Btn onClick={exportLedgerExcel} variant="secondary" size="sm" icon="⬇">Excel</Btn>
        <Btn onClick={exportLedgerPDF} variant="secondary" size="sm" icon="⬇">PDF</Btn>
      </div>
    </div>

    {metrics&&<div style={{display:'grid',gridTemplateColumns:mobile?'repeat(2,1fr)':isSupplier?'repeat(3,1fr)':'repeat(4,1fr)',gap:8,marginBottom:16}}>
      {!isSupplier&&<div style={{background:'#F9FAFB',borderRadius:7,padding:'8px 10px',border:`1px solid ${C.bdr}`}}>
        <div style={{fontSize:9,color:C.sub,textTransform:'uppercase',letterSpacing:.4}}>Profit</div>
        <div style={{fontSize:14,fontWeight:800,color:metrics.profit>=0?C.blue:C.red}}>{fmt(metrics.profit)}</div>
      </div>}
      {!isSupplier&&<div style={{background:'#F9FAFB',borderRadius:7,padding:'8px 10px',border:`1px solid ${C.bdr}`}}>
        <div style={{fontSize:9,color:C.sub,textTransform:'uppercase',letterSpacing:.4}}>Margin</div>
        <div style={{fontSize:14,fontWeight:800,color:C.txt}}>{metrics.avgMarginPct.toFixed(1)}%</div>
      </div>}
      <div style={{background:'#F9FAFB',borderRadius:7,padding:'8px 10px',border:`1px solid ${C.bdr}`}}>
        <div style={{fontSize:9,color:C.sub,textTransform:'uppercase',letterSpacing:.4}}>Avg Payment Days</div>
        <div style={{fontSize:14,fontWeight:800,color:C.txt}}>{metrics.avgPaymentDays===null?'—':metrics.avgPaymentDays+'d'}</div>
      </div>
      <div style={{background:'#F9FAFB',borderRadius:7,padding:'8px 10px',border:`1px solid ${C.bdr}`}}>
        <div style={{fontSize:9,color:C.sub,textTransform:'uppercase',letterSpacing:.4}}>Grade</div>
        <div style={{fontSize:14,fontWeight:800,color:metrics.grade?gradeColor[metrics.grade]:C.muted}}>{metrics.grade||'—'}</div>
      </div>
    </div>}

    {!isSupplier&&!isTransporterParty&&<div style={{marginBottom:16,background:'#FAF5FF',borderRadius:8,border:`1px solid ${C.purple}30`,padding:'10px 14px'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
        <div>
          <div style={{fontSize:10,fontWeight:700,color:C.purple,textTransform:'uppercase',letterSpacing:.5}}>Repair Expenditure</div>
          <div style={{fontSize:16,fontWeight:800,color:C.purple}}>{fmt(repairExpTotal)}{repairExps.length>0&&<span style={{fontSize:11,fontWeight:400,color:C.sub,marginLeft:6}}>{repairExps.length} entr{repairExps.length===1?'y':'ies'}</span>}</div>
        </div>
        <Btn onClick={()=>setShowRepairModal(true)} variant="secondary" size="sm">+ Repair Expenditure</Btn>
      </div>
      {repairExps.length>0&&<div style={{marginTop:8,borderTop:`1px solid ${C.purple}20`,paddingTop:8}}>
        {repairExps.slice(0,5).map(r=><div key={r.id} style={{display:'flex',justifyContent:'space-between',fontSize:11,color:C.sub,padding:'2px 0'}}>
          <span>{fmtD(r.date)} · {r.itemName||(r.source==='payment'?'Reimbursement':'—')}{r.notes?' · '+r.notes:''}</span>
          <span style={{fontWeight:700,color:C.purple}}>{fmt(r.expenseAmount)}</span>
        </div>)}
        {repairExps.length>5&&<div style={{fontSize:10,color:C.muted,marginTop:2}}>+{repairExps.length-5} more</div>}
      </div>}
    </div>}

    <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>Recent Activity {rows.length>0&&`(${rows.length})`}</div>
    <div style={{border:`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden',marginBottom:8,background:mobile?'transparent':undefined,padding:mobile?8:0}}>
      <CardTable mobile={mobile} rows={actPageRows} keyFn={(r,i)=>i} emptyLabel="No sales or payments recorded yet"
        onRowClick={r=>{
          if(r.type==='sale'&&r.saleId){const found=state.sales.find(x=>x.id===r.saleId);if(found)setViewSale(found);}
          else if(r.type==='purchase'&&r.purchaseId){const found=state.purchases.find(x=>x.id===r.purchaseId);if(found)setViewPurchase(found);}
        }}
        footer={<div style={{display:'flex',justifyContent:'space-between',padding:'9px 12px',background:'#F6F8FA',fontWeight:700,color:C.txt}}>
          <span>Opening Balance</span><span style={{fontWeight:800,color:(party.ob||0)>0?C.red:C.green}}>{fmt(party.ob||0)}</span>
        </div>}
        columns={[
          {key:'date',label:'Date',render:r=><span style={{color:C.sub,whiteSpace:'nowrap'}}>{fmtD(r.date)}</span>},
          {key:'type',label:'Type',render:r=><Badge v={typeMeta[r.type].label} color={typeMeta[r.type].badge}/>},
          {key:'desc',label:'Description',render:r=>{
            const clickable=(r.type==='sale'&&r.saleId)||(r.type==='purchase'&&r.purchaseId);
            return <span style={{fontSize:12,color:clickable?C.blue:C.sub,textDecoration:clickable?'underline':'none',textDecorationStyle:'dotted'}}>{r.desc}</span>;
          }},
          {key:'amount',label:'Amount',align:'right',render:r=><span style={{fontWeight:700,color:typeMeta[r.type].color,fontVariantNumeric:'tabular-nums'}}>{r.delta>0?'+':r.delta<0?'−':''}{fmt(Math.abs(r.delta))}</span>},
          {key:'balance',label:'Balance',align:'right',render:r=><span style={{fontWeight:700,color:r.balance>0?C.red:C.green,fontVariantNumeric:'tabular-nums'}}>{fmt(r.balance)}</span>},
        ]}/>
    </div>
    {actTotalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginBottom:16}}>
      <Btn onClick={()=>setActPage(p=>Math.max(0,p-1))} disabled={actPage===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {actPage+1} of {actTotalPages}</span>
      <Btn onClick={()=>setActPage(p=>Math.min(actTotalPages-1,p+1))} disabled={actPage>=actTotalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}

    {histRows.length>0&&<div>
      <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>History — rolled into opening balance ({histRows.length})</div>
      <div style={{border:`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden',background:mobile?'transparent':undefined,padding:mobile?8:0}}>
        <CardTable mobile={mobile} rows={histPageRows} keyFn={(r,i)=>i}
          onRowClick={r=>{
            const clickable=!isSupplier&&!isTransporterParty&&r.type==='sale'&&r.items;
            if(!clickable)return;
            setViewSale({id:null,date:r.date,partyId:party.id,partyName:party.name,total:r.amount,creditDays:party.credit||30,
              items:r.items.map(x=>Array.isArray(x)?{itemId:null,itemName:x[0],customerName:x[0],qty:x[1],rate:x[2],amount:x[1]*x[2]}:x)});
          }}
          columns={[
            {key:'date',label:'Date',render:r=><span style={{color:C.sub,whiteSpace:'nowrap'}}>{fmtD(r.date)}</span>},
            {key:'type',label:'Type',render:r=><Badge v={typeMeta[r.type].label} color={typeMeta[r.type].badge}/>},
            {key:'desc',label:'Description',render:r=>{
              const clickable=!isSupplier&&!isTransporterParty&&r.type==='sale'&&r.items;
              return <span style={{fontSize:12,color:clickable?C.blue:C.sub,textDecoration:clickable?'underline':'none',textDecorationStyle:'dotted'}}>{r.desc}</span>;
            }},
            {key:'amount',label:'Amount',align:'right',render:r=><span style={{fontWeight:600,color:typeMeta[r.type].color,fontVariantNumeric:'tabular-nums'}}>{fmt(r.amount)}</span>},
            {key:'balance',label:'Balance',align:'right',render:r=><span style={{fontWeight:700,color:r.balance>0?C.red:C.green,fontVariantNumeric:'tabular-nums'}}>{fmt(r.balance)}</span>},
          ]}/>
      </div>
      {histTotalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:8}}>
        <Btn onClick={()=>setHistPage(p=>Math.max(0,p-1))} disabled={histPage===0} variant="secondary" size="sm">← Prev</Btn>
        <span style={{fontSize:12,color:C.sub}}>Page {histPage+1} of {histTotalPages}</span>
        <Btn onClick={()=>setHistPage(p=>Math.min(histTotalPages-1,p+1))} disabled={histPage>=histTotalPages-1} variant="secondary" size="sm">Next →</Btn>
      </div>}
    </div>}
  </Modal>;
}

// Standalone repair-expenditure log: a warranty part (PCB) was handed to this
// party — sometimes billed, sometimes free — and the faulty one has come
// back. Books the net cost (part cost minus whatever was billed) against the
// party for reporting, and decrements the part's stock since it physically
// left the shop. Deliberately doesn't touch the party's receivable — if the
// replacement was billed, that's recorded as an ordinary sale separately;
// reimbursing a party for repairs THEY paid for uses the "Repair ₹" field in
// Accounts/Outstanding instead, which does adjust what they owe.
function RepairExpenseModal({party,onClose,onSaved}){
  const[date,setDate]=useState(today());
  const[iq,setIq]=useState('');
  const[itemId,setItemId]=useState('');
  const[qty,setQty]=useState('1');
  const[billed,setBilled]=useState('0');
  const[expense,setExpense]=useState('');
  const[expTouched,setExpTouched]=useState(false);
  const[notes,setNotes]=useState('');
  const{toast,TC}=useToast();
  const candidates=useMemo(()=>iq.length<1?[]:state.items.filter(it=>it.active!==false&&(it.internal||it.name||'').toLowerCase().includes(iq.toLowerCase())).slice(0,8),[iq]);
  const selItem=state.items.find(it=>it.id===itemId);

  useEffect(()=>{
    if(expTouched)return;
    const q=Number(qty)||0;
    const cost=selItem?(selItem.purch||0)*q:0;
    const b=Number(billed)||0;
    setExpense(String(Math.max(0,Math.round(cost-b))));
  },[selItem,qty,billed,expTouched]);

  const doSave=()=>{
    const q=Number(qty)||0;
    const expAmt=Number(expense)||0;
    if(!selItem)return toast('Pick the PCB/part given','error');
    if(q<=0)return toast('Enter a valid quantity','error');
    if(expAmt<=0)return toast('Enter an expense amount','error');
    adjustStock(selItem.id,-q);
    recordRepairExpense({partyId:party.id,partyName:party.name,date,itemId:selItem.id,itemName:selItem.internal||selItem.name,qty:q,billedAmount:Number(billed)||0,expenseAmount:expAmt,source:'return',notes});
    toast('Repair expenditure recorded — '+fmt(expAmt)+' ✓');
    onSaved?.();
  };

  return <Modal title={'Record Repair Expenditure — '+party.name} onClose={onClose} width={460}>
    <TC/>
    <div style={{fontSize:12,color:C.sub,marginBottom:14}}>Logs the cost of a warranty part given to this party and decrements its stock — doesn't change what they owe. If it was billed, record that separately as a normal sale; use the "Repair ₹" field at payment time instead to reimburse a party for repairs they paid for themselves.</div>
    <Grid2>
      <Inp label="Date" value={date} onChange={setDate} type="date" required/>
      <Inp label="Qty" value={qty} onChange={setQty} type="number"/>
      <div style={{gridColumn:'1/-1'}}>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>Part / PCB</label>
        {selItem
          ?<div style={{display:'flex',alignItems:'center',gap:8,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
            <span style={{fontWeight:700,flex:1,fontSize:13,color:C.txt}}>{selItem.internal||selItem.name}</span>
            <span style={{fontSize:11,color:C.sub}}>{selItem.stock||0} pcs in stock</span>
            <button onClick={()=>{setItemId('');setIq('');}} style={{background:'none',border:'none',color:C.blue,cursor:'pointer',fontSize:11,textDecoration:'underline',fontFamily:'inherit'}}>Change</button>
          </div>
          :<div style={{position:'relative'}}>
            <input value={iq} onChange={e=>setIq(e.target.value)} placeholder="Search item…" autoFocus autoComplete="off"
              style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
              onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
            {candidates.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 8px 8px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)'}}>
              {candidates.map(it=><div key={it.id} onClick={()=>{setItemId(it.id);setIq('');}} style={{padding:'8px 12px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`}}
                onMouseEnter={e=>e.currentTarget.style.background='#F9FAFB'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                <div style={{fontWeight:600,color:C.txt}}>{it.internal||it.name}</div>
                <div style={{fontSize:11,color:C.sub}}>{it.cat} · {it.stock||0} pcs in stock · cost {fmt(it.purch||0)}</div>
              </div>)}
            </div>}
          </div>}
      </div>
      <Inp label="Billed to party ₹ (0 if free)" value={billed} onChange={setBilled} type="number"/>
      <Inp label="Expense to book ₹" value={expense} onChange={v=>{setExpense(v);setExpTouched(true);}} type="number"/>
      <div style={{gridColumn:'1/-1'}}><Inp label="Notes (optional)" value={notes} onChange={setNotes} placeholder="e.g. PCB replaced under warranty"/></div>
    </Grid2>
    <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:16}}>
      <Btn onClick={onClose} variant="secondary">Cancel</Btn>
      <Btn onClick={doSave}>Record</Btn>
    </div>
  </Modal>;
}

// ── ITEM MASTERS ─────────────────────────────────────────────────────────────

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
// ── ACCOUNTS MODULE ───────────────────────────────────────────────────────────


// Quick-add party mini modal
function QuickAddPartyModal({defaultName, type='customer', onAdded, onClose}) {
  const [name,setName] = useState(defaultName||'');
  const [city,setCity] = useState('');
  const [route,setRoute] = useState('Raipur');
  const [credit,setCredit] = useState(30);
  const {toast,TC} = useToast();
  const save_party = () => {
    if (!name.trim()) return toast('Name required','error');
    const p = {id:uid('PTY'),name:name.trim(),city,route,credit:Number(credit)||30,ob:0,type,active:true};
    state.parties = [...state.parties, p];
    save.parties();
    toast((type==='supplier'?'Supplier':type==='transporter'?'Transporter':'Party')+' added ✓');
    onAdded(p);
  };
  return <Modal title={type==='supplier'?'Quick Add Supplier':type==='transporter'?'Quick Add Transporter':'Quick Add Party'} onClose={onClose} width={420}>
    <TC/>
    <div style={{display:'flex',flexDirection:'column',gap:12}}>
      <Inp label="Party Name" value={name} onChange={setName} required autoFocus/>
      <Grid2>
        <Inp label="City" value={city} onChange={setCity} placeholder="e.g. Raipur"/>
        <Sel label="Route" value={route} onChange={setRoute} options={[...ROUTES,'Other']}/>
      </Grid2>
      <Inp label="Credit Days" value={credit} onChange={v=>setCredit(Number(v)||30)} type="number"/>
      <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:4}}>
        <Btn onClick={onClose} variant="secondary" size="sm">Cancel</Btn>
        <Btn onClick={save_party} size="sm">Add Party</Btn>
      </div>
    </div>
  </Modal>;
}

function getAllCats() { return [...TXN_CATEGORIES, ...state.customCategories]; }
function getCatLabel(id) { return getAllCats().find(c=>c.id===id)?.label || id || '—'; }

function NewCategoryModal({onClose,onCreated}) {
  const [label,setLabel]=useState('');
  const [dir,setDir]=useState('out');
  const {toast,TC}=useToast();
  const doSave=()=>{
    if(!label.trim())return toast('Name required','error');
    const cat={id:'cc_'+uid(''),label:label.trim(),dir,party:null,icon:dir==='in'?'↑':'↓'};
    state.customCategories=[...state.customCategories,cat]; save.customCategories();
    toast('Category added ✓'); onCreated(cat);
  };
  return <Modal title="New Category" onClose={onClose} width={380}>
    <TC/>
    <div style={{display:'flex',flexDirection:'column',gap:12}}>
      <Inp label="Category Name" value={label} onChange={setLabel} placeholder="e.g. Repairs & Maintenance" required autoFocus/>
      <div>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:4}}>Default Direction</label>
        <div style={{display:'flex',gap:6}}>
          {['in','out'].map(d=><button key={d} onClick={()=>setDir(d)} style={{flex:1,padding:'8px 10px',borderRadius:7,border:`1.5px solid ${dir===d?(d==='in'?C.green:C.red):C.bdr}`,background:dir===d?(d==='in'?C.greenSoft:C.redSoft):C.surf,color:dir===d?(d==='in'?C.green:C.red):C.sub,fontSize:13,fontWeight:700,cursor:'pointer',fontFamily:'inherit'}}>{d==='in'?'↑ Money In':'↓ Money Out'}</button>)}
        </div>
      </div>
      <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:4}}>
        <Btn onClick={onClose} variant="secondary" size="sm">Cancel</Btn>
        <Btn onClick={doSave} size="sm">Add Category</Btn>
      </div>
    </div>
  </Modal>;
}

function EditTxnModal({txn,onClose,onSaved}) {
  const allCats=getAllCats();
  const origMeta = txn.type==='transfer' ? {id:'transfer',label:'Transfer',dir:'contra',party:null}
    : txn.type==='party_payment' ? {id:'party_receipt',label:'Party Payment / Receipt',dir:'in',party:'customer'}
    : txn.type==='receipt' ? {id:'other_receipt',label:'Other Receipt',dir:'in',party:null}
    : (allCats.find(c=>c.id===txn.category) || {id:txn.category,label:getCatLabel(txn.category),dir:'out',party:null});
  const lockCategory = !!txn.partyId;
  const candidateCats = lockCategory ? [origMeta] : allCats.filter(c=>c.party===null&&c.id!=='transfer');
  const [categoryId,setCategoryId]=useState(origMeta.id);
  const catMeta = allCats.find(c=>c.id===categoryId) || origMeta;
  const initDir = origMeta.dir==='contra' ? 'contra' : (txn.amount>=0 ? origMeta.dir : (origMeta.dir==='in'?'out':'in'));
  const [dir,setDir]=useState(initDir);
  const [amount,setAmount]=useState(String(Math.abs(txn.amount)));
  const [date,setDate]=useState(txn.date);
  const [notes,setNotes]=useState(txn.notes||'');
  const {toast,TC}=useToast();

  const doSave=()=>{
    const amt=parseFloat(amount);
    if(!amt||amt<=0) return toast('Enter a valid amount','error');
    const signedAmt = catMeta.dir==='contra' ? amt : (dir===catMeta.dir?amt:-amt);
    let savedType=txn.type, savedCategory=txn.category, savedSource=txn.source;
    if(!lockCategory){
      if(categoryId==='other_receipt'){ savedType='receipt'; savedCategory=undefined; savedSource='other'; }
      else { savedType='payment'; savedCategory=categoryId; savedSource=undefined; }
    }
    state.txns = state.txns.map(t=>t.id===txn.id?{...t,date,amount:signedAmt,notes,type:savedType,category:savedCategory,source:savedSource}:t);
    save.txns();
    state.payments = state.payments.map(p=>p.txnId===txn.id?{...p,date,amount:signedAmt,notes}:p);
    save.payments();
    state.supplierPayments = state.supplierPayments.map(p=>p.txnId===txn.id?{...p,date,amount:signedAmt,notes}:p);
    save.supplierPayments();
    toast('Updated ✓'); onSaved?.(); onClose();
  };

  return <Modal title="Edit Transaction" onClose={onClose} width={420}>
    <TC/>
    <div style={{display:'flex',flexDirection:'column',gap:12}}>
      <Inp label="Date" value={date} onChange={setDate} type="date"/>
      {lockCategory
        ? <div style={{fontSize:12,color:C.sub,background:'#F9FAFB',padding:'8px 12px',borderRadius:7,border:`1px solid ${C.bdr}`}}>
            Category: <b style={{color:C.txt}}>{origMeta.label}</b>{txn.partyName?<> — linked to <b style={{color:C.txt}}>{txn.partyName}</b></>:''}
            <div style={{fontSize:11,color:C.muted,marginTop:2}}>Party-linked entries keep their category — delete and re-enter to change the party.</div>
          </div>
        : <div>
            <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:4}}>Category</label>
            <select value={categoryId} onChange={e=>{setCategoryId(e.target.value);const m=allCats.find(c=>c.id===e.target.value);setDir(m?.dir||'out');}}
              style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,fontFamily:'inherit',color:C.txt,boxSizing:'border-box'}}>
              {candidateCats.map(c=><option key={c.id} value={c.id}>{c.label}</option>)}
            </select>
          </div>}
      {catMeta.dir!=='contra'&&<div>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:4}}>Direction</label>
        <div style={{display:'flex',gap:6}}>
          {['in','out'].map(d=><button key={d} onClick={()=>setDir(d)} style={{flex:1,padding:'7px 10px',borderRadius:7,border:`1.5px solid ${dir===d?(d==='in'?C.green:C.red):C.bdr}`,background:dir===d?(d==='in'?C.greenSoft:C.redSoft):C.surf,color:dir===d?(d==='in'?C.green:C.red):C.sub,fontSize:12,fontWeight:700,cursor:'pointer',fontFamily:'inherit'}}>{d==='in'?'↑ Money In':'↓ Money Out'}</button>)}
        </div>
      </div>}
      <Inp label="Amount ₹" value={amount} onChange={setAmount} type="number"/>
      <Inp label="Notes" value={notes} onChange={setNotes}/>
      <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:4}}>
        <Btn onClick={onClose} variant="secondary" size="sm">Cancel</Btn>
        <Btn onClick={doSave} size="sm">Save Changes</Btn>
      </div>
    </div>
  </Modal>;
}

function AccountsModule() {
  // ── FORM STATE — category/account/date persist between saves ──────────────
  const [categoryId,setCategoryId] = useState('party_receipt');
  const [dir,setDir]         = useState('in');
  const [date,setDate]       = useState(today());
  const [account,setAccount] = useState('cash');
  const [toAccount,setToAcc] = useState('upkaran');
  // ── cleared on each save ─────────────────────────────────────────────────
  const [amount,setAmount]   = useState('');
  const [payee,setPayee]     = useState('');
  const [partyId,setPId]     = useState('');
  const [pq,setPq]           = useState('');
  const [pIdx,setPIdx]       = useState(0);
  const [notes,setNotes]     = useState('');
  const [selInvIds,setSelInvIds] = useState([]);   // selected invoice dates for this payment
  const [discount,setDiscount]   = useState({});   // per-invoice discount {disc_DATE: amount}
  const [supplierDiscount,setSupplierDiscount] = useState('');   // aggregate write-off against a vendor payment
  // Repair expenditure: same per-invoice mechanic as discount (reduces the
  // bill), but reimburses a customer for warranty repairs they paid for
  // themselves rather than forgiving old debt — so it also books a real
  // expense (repairExpOv lets that expense amount be edited separately from
  // what's deducted from the bill, defaulting to match it).
  const [repair,setRepair]       = useState({});   // per-invoice {rep_DATE: amount deducted from bill}
  const [repairExpOv,setRepairExpOv] = useState({}); // per-invoice {rep_DATE: expense override, defaults to repair amount}
  // ── UI state ─────────────────────────────────────────────────────────────
  const [ledger,setLedger]   = useState(null);
  const [quickAdd,setQuickAdd] = useState(null);
  const [newCat,setNewCat]   = useState(false);
  const [editTxn,setEditTxn] = useState(null);
  const [showReturn,setShowReturn] = useState(false);
  const [editReturnSale,setEditReturnSale] = useState(null);
  const [catLedger,setCatLedger] = useState(null);
  const [catFilter,setCatFilter] = useState('');
  const [srch,setSrch]       = useState('');
  const [page,setPage]       = useState(0);
  const [_fvAcc, _setFvAcc] = useState(0);
  const forceUpdate = () => _setFvAcc(v => v+1);
  const PAGE = 30;
  const {toast,TC} = useToast();
  const mobile = useIsMobile();

  const allCats = useMemo(()=>getAllCats(),[state.customCategories.length]);
  const catMeta = useMemo(()=>allCats.find(c=>c.id===categoryId)||allCats[0],[allCats,categoryId]);
  const isTransfer  = categoryId==='transfer';
  const partyKind    = catMeta.party;
  const isSupplier    = partyKind==='supplier';
  const isTransporter = partyKind==='transporter';
  const isCustomerParty = partyKind==='customer';
  const isParty  = !!partyKind;

  const chooseCategory = id => {
    if(id==='__new__'){ setNewCat(true); return; }
    setCategoryId(id);
    const m=allCats.find(c=>c.id===id);
    setDir(m?.dir==='contra'?'contra':(m?.dir||'out'));
    setPId(''); setPq(''); setPayee(''); setSelInvIds([]); setDiscount({}); setSupplierDiscount(''); setRepair({}); setRepairExpOv({});
    // Switching into Transfer with From/To already pointing at the same
    // account (e.g. left over from before this category was picked) would
    // hit the same "From and To must differ" trap fixed on the From-Account
    // select below — same fix, applied at the other entry point.
    if(id==='transfer'&&account===toAccount){
      const fallback=ACCOUNTS.find(a=>a.id!==account);
      if(fallback)setToAcc(fallback.id);
    }
  };
  const categoryTypeahead = useCategoryTypeahead(allCats, chooseCategory);
  const catFilterTypeahead = useCategoryTypeahead(allCats, id=>{setCatFilter(id);setPage(0);});

  const custs    = useMemo(()=>state.parties.filter(p=>p.type==='customer'),[state.parties.length]);
  const supps    = useMemo(()=>state.parties.filter(p=>p.type==='supplier'),[state.parties.length]);
  const transporters = useMemo(()=>state.parties.filter(p=>p.type==='transporter'),[state.parties.length]);
  const partyList = isSupplier?supps:isTransporter?transporters:custs;
  const selParty = useMemo(()=>partyList.find(p=>p.id===partyId),[partyList,partyId]);
  const partyOD  = useMemo(()=>{
    if(!selParty) return 0;
    if(isSupplier||isTransporter) return getSupplierBal(selParty.id);
    return getCustomerBal(selParty.id);
  },[selParty?.id,isSupplier,isTransporter,state.sales.length,state.payments.length,state.txns.length,state.purchases.length,state.supplierPayments.length,state.freightCharges.length]);
  const partyBills = useMemo(()=>{
    if(!selParty) return [];
    if(isSupplier||isTransporter) return getSupplierBills(selParty);
    return getPartyBills(selParty);
  },[selParty?.id,isSupplier,isTransporter,state.sales.length,state.payments.length,state.purchases.length,state.freightCharges.length,state.supplierPayments.length]);
  // Returns recorded for this party, so a mis-entered return (wrong invoice,
  // wrong qty) can be corrected right here instead of hunting for it in the
  // Sales Log — surfaced wherever it's created (via "+ Add Return" above).
  const partyReturns = useMemo(()=>{
    if(!selParty||isSupplier||isTransporter) return [];
    return state.sales.filter(s=>s.partyId===selParty.id&&s.type==='return').sort((a,b)=>b.date.localeCompare(a.date)||(b.createdAt||'').localeCompare(a.createdAt||''));
  },[selParty?.id,isSupplier,isTransporter,state.sales.length]);
  const psugg    = useMemo(()=>pq.length<1?[]:partyList.filter(p=>p.name.toLowerCase().includes(pq.toLowerCase())||p.city?.toLowerCase().includes(pq.toLowerCase())).slice(0,7),[partyList,pq]);
  // Non-party categories (Salary, Ghar Kharch, Office Expenses, etc.) have no
  // party master to pick from, but the same names recur constantly — e.g.
  // the same 4-5 staff names every month for Salary, or "Milk"/"Grocery"/
  // "Maid" for Ghar Kharch. Build a per-category pick-list from this app's
  // own past entries (ranked by how often used, then most recent) so typing
  // one letter surfaces it instead of retyping the full name every time.
  const payeeHistory = useMemo(()=>{
    if(isParty) return [];
    const counts=new Map();
    for(const t of state.txns){
      if(t.category!==categoryId) continue;
      const name=(t.payee||'').trim();
      if(!name) continue;
      const key=name.toLowerCase();
      const cur=counts.get(key);
      if(cur){ cur.count++; if(t.date>cur.lastDate) cur.lastDate=t.date; }
      else counts.set(key,{name,count:1,lastDate:t.date});
    }
    return [...counts.values()].sort((a,b)=>b.count-a.count||b.lastDate.localeCompare(a.lastDate));
  },[categoryId,isParty,state.txns.length]);
  const payeeSugg = useMemo(()=>{
    if(isParty) return [];
    const q=payee.trim().toLowerCase();
    return (q?payeeHistory.filter(p=>p.name.toLowerCase().includes(q)):payeeHistory).slice(0,7);
  },[payeeHistory,payee,isParty]);
  // A discount/write-off can be recorded on its own, with no cash actually
  // changing hands — e.g. writing off the last ₹630 of an old invoice — so
  // Save must accept that case even though the Amount field is 0.
  const discTotal=Object.values(discount||{}).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const suppDiscTotal=parseFloat(supplierDiscount)||0;
  const repairTotal=Object.values(repair||{}).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const canSave=(parseFloat(amount)||0)>0||(isCustomerParty&&discTotal>0)||(isSupplier&&suppDiscTotal>0)||(isCustomerParty&&repairTotal>0);
  const showAdd  = pq.length>=2&&psugg.length===0;

  const balances = useMemo(()=>ACCOUNTS.map(a=>({...a,bal:getAccBal(a.id)})),[state.txns.length,state.payments.length]);
  const totalBal = balances.reduce((s,a)=>s+a.bal,0);
  const allTxns  = useMemo(()=>[...state.txns].sort((a,b)=>b.date.localeCompare(a.date)||((b.createdAt||'').localeCompare(a.createdAt||''))),[state.txns.length,_fvAcc]);
  const matchesCat = (t,cid) => {
    if(!cid) return true;
    if(cid==='party_receipt') return t.type==='party_payment';
    if(cid==='other_receipt') return t.type==='receipt';
    if(cid==='transfer') return t.type==='transfer';
    return t.category===cid;
  };
  const filtered = useMemo(()=>allTxns.filter(t=>matchesCat(t,catFilter)&&(!srch||(t.payee||'').toLowerCase().includes(srch.toLowerCase())||(t.partyName||'').toLowerCase().includes(srch.toLowerCase()))),[allTxns,catFilter,srch]);
  const catTotals = useMemo(()=>{
    if(!catFilter) return null;
    let inSum=0,outSum=0;
    for(const t of filtered){
      const out=(t.type==='payment')===(t.amount>=0);
      if(out) outSum+=Math.abs(t.amount); else inSum+=Math.abs(t.amount);
    }
    return {inSum,outSum,net:inSum-outSum};
  },[filtered,catFilter]);
  const pageData = filtered.slice(page*PAGE,(page+1)*PAGE);
  const totalPages = Math.ceil(filtered.length/PAGE);

  // Focus helpers
  const foc = id => setTimeout(()=>{const el=document.getElementById(id);if(el){el.focus();if(el.tagName==='INPUT')el.select();}},40);
  useEffect(()=>{foc('acct-date');},[]);

  // Advancing date with arrow keys - handled natively by browser's date input
  // Only intercept Enter to advance to next field
  const dateKey = e => {
    if(e.key==='Enter'){e.preventDefault();nextFocus('date');}
  };

  // Cursor flow: date → category → party/payee → amount → account → notes → save
  const FIELD_ORDER = ['date','category','party','amount','account','notes'];
  const nextFocus = cur => {
    const idx = FIELD_ORDER.indexOf(cur);
    const next = FIELD_ORDER[idx+1];
    if(next) foc('acct-'+next);
  };

  const enterFrom = field => e => {
    if(e.key==='Enter'){e.preventDefault();nextFocus(field);}
  };
  const selKey = field => e => {
    if(e.key==='Enter'){e.preventDefault();nextFocus(field);}
  };

  // Clear transient fields on save (keep date/category/account) — direction is
  // NOT kept: it always resets to the category's default so a manual override
  // used for one entry never silently carries over to the next.
  const clearTransient = () => {
    setAmount('');setPayee('');setPId('');setPq('');setPIdx(0);setNotes('');setSelInvIds([]);setDiscount({});setSupplierDiscount('');setRepair({});setRepairExpOv({});
    setDir(catMeta.dir==='contra'?'contra':(catMeta.dir||'out'));
    foc('acct-party');
  };

  const doSave = () => {
    const amtRaw=parseFloat(amount)||0;
    if(!canSave) return toast('Enter a valid amount or discount','error');
    if(isTransfer&&account===toAccount) return toast('From and To must differ','error');
    // A party-type category (supplier/customer/transporter payment) with text
    // typed into the party box but never actually selected from the dropdown
    // leaves partyId empty. Every write below is gated on `&&partyId`, so
    // without this check the generic txn still saves and shows "Saved ✓" —
    // silently dropping the party attribution instead of erroring. Block it
    // here instead, with an explicit message, so it can't happen unnoticed.
    if(isParty&&!partyId) return toast('Select the party from the dropdown before saving','error');

    // The cash-movement side is entirely optional — a discount/write-off can
    // stand alone with no payment attached (amtRaw===0), e.g. writing off the
    // last ₹630 of an old invoice with nothing actually paid.
    if(amtRaw>0){
      const signedAmt = isTransfer ? amtRaw : (dir===catMeta.dir?amtRaw:-amtRaw);

      let savedType,savedCategory,savedSource;
      if(isTransfer) savedType='transfer';
      else if(categoryId==='party_receipt') savedType='party_payment';
      else if(categoryId==='other_receipt'){ savedType='receipt'; savedSource='other'; }
      else { savedType='payment'; savedCategory=categoryId; }

      const t={
        id:uid('TXN'),date,type:savedType,account,
        toAccount:isTransfer?toAccount:undefined,
        amount:signedAmt,
        payee:(savedType==='party_payment'||isSupplier||isTransporter)?(selParty?.name||''):(payee||''),
        partyId:isParty?partyId:undefined,
        partyName:selParty?.name||undefined,
        category:savedCategory,
        source:savedSource,
        notes,createdAt:new Date().toISOString(),
      };
      state.txns=[t,...state.txns]; save.txns();
      if((isSupplier||isTransporter)&&partyId){
        const sp={id:uid('SPMT'),txnId:t.id,date,partyId,partyName:selParty?.name||'',amount:signedAmt,account,notes};
        state.supplierPayments=[sp,...state.supplierPayments]; save.supplierPayments();
      } else if(isParty&&partyId){
        const pmt={id:uid('PMT'),txnId:t.id,date,partyId,partyName:selParty?.name||'',amount:signedAmt,account,notes};
        state.payments=[pmt,...state.payments]; save.payments();
      }
    }
    // Also save discount/write-off if entered (customer invoices only)
    const disc=isCustomerParty?discTotal:0;
    if(disc>0&&partyId){
      // The discount state keys are 'disc_'+bill.key, and bill.key is itself
      // '(N|H)_<date>_<index>' (see getPartyBills) — pull just the date back
      // out rather than feeding the whole key to fmtD, which silently renders
      // as "Invalid Date" for anything that isn't a bare YYYY-MM-DD string.
      const wo={id:uid('WO'),type:'writeoff',date,partyId,partyName:selParty?.name||'',amount:disc,account:'writeoff',notes:'Discount/write-off '+Object.entries(discount||{}).filter(([,v])=>parseFloat(v)>0).map(([k,v])=>{const m=k.match(/\d{4}-\d{2}-\d{2}/);return(m?fmtD(m[0]):'')+' ₹'+v;}).join(', '),createdAt:new Date().toISOString()};
      state.txns=[wo,...state.txns]; save.txns();
      const woPmt={id:uid('WOP'),txnId:wo.id,date,partyId,partyName:selParty?.name||'',amount:disc,account:'writeoff',notes:'Discount write-off'};
      state.payments=[woPmt,...state.payments]; save.payments();
    }
    // Repair expenditure: reduces the bill exactly like a discount (same
    // 'writeoff' account, no real cash involved) but additionally books a
    // repairExpenses record per invoice — kept separate from the discount
    // write-off above so P&L can report warranty cost as its own line.
    if(isCustomerParty&&repairTotal>0&&partyId){
      const wo={id:uid('WO'),type:'writeoff',date,partyId,partyName:selParty?.name||'',amount:repairTotal,account:'writeoff',notes:'Repair expenditure '+Object.entries(repair||{}).filter(([,v])=>parseFloat(v)>0).map(([k,v])=>{const m=k.match(/\d{4}-\d{2}-\d{2}/);return(m?fmtD(m[0]):'')+' ₹'+v;}).join(', '),createdAt:new Date().toISOString()};
      state.txns=[wo,...state.txns]; save.txns();
      const woPmt={id:uid('WOP'),txnId:wo.id,date,partyId,partyName:selParty?.name||'',amount:repairTotal,account:'writeoff',notes:'Repair expenditure adjustment'};
      state.payments=[woPmt,...state.payments]; save.payments();
      Object.entries(repair||{}).forEach(([k,v])=>{
        const amt=parseFloat(v)||0;
        if(amt<=0)return;
        const m=k.match(/\d{4}-\d{2}-\d{2}/);
        const expAmt=parseFloat(repairExpOv[k])||amt;
        recordRepairExpense({partyId,partyName:selParty?.name||'',date,itemId:'',itemName:'',qty:0,billedAmount:0,expenseAmount:expAmt,source:'payment',linkedTxnId:wo.id,notes:'Reimbursed against '+(m?fmtD(m[0]):'')+' invoice'});
      });
    }
    // Purchase-side write-off (reduces what we owe a supplier — no per-invoice
    // breakdown exists for purchases, so this is a single aggregate amount
    // rather than the per-invoice discount customers get).
    const suppDisc=isSupplier?suppDiscTotal:0;
    if(suppDisc>0&&partyId){
      const wo={id:uid('WO'),type:'writeoff',date,partyId,partyName:selParty?.name||'',amount:suppDisc,account:'writeoff',notes:'Purchase discount/write-off',createdAt:new Date().toISOString()};
      state.txns=[wo,...state.txns]; save.txns();
      const woSp={id:uid('SPMT'),txnId:wo.id,date,partyId,partyName:selParty?.name||'',amount:suppDisc,account:'writeoff',notes:'Purchase discount write-off'};
      state.supplierPayments=[woSp,...state.supplierPayments]; save.supplierPayments();
    }
    toast('Saved ✓'); forceUpdate(n=>n+1); clearTransient();
  };

  const deleteTxn = id => {
    if(!window.confirm('Delete?')) return;
    state.txns=state.txns.filter(t=>t.id!==id); save.txns();
    state.payments=state.payments.filter(p=>p.txnId!==id); save.payments();
    state.supplierPayments=state.supplierPayments.filter(p=>p.txnId!==id); save.supplierPayments();
    forceUpdate(n=>n+1);
  };

  const TYPE_LABEL={receipt:'Receipt',payment:'Expense',transfer:'Transfer',party_payment:'Party Pmt'};
  const TYPE_COLOR={receipt:C.green,payment:C.red,transfer:C.blue,party_payment:C.green};

  // Shared input style
  const fieldStyle = (active,extra={}) => ({
    width:'100%',padding:'9px 11px',borderRadius:7,fontFamily:'inherit',fontSize:13,outline:'none',
    border:`1.5px solid ${active?C.blue:C.bdr}`,background:C.surf,boxSizing:'border-box',...extra
  });
  const labelStyle = {fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.6,display:'block',marginBottom:4};

  const [focusedField,setFocused] = useState('date');

  return <div>
    <TC/>
    {ledger&&<AccountLedger accId={ledger} onClose={()=>setLedger(null)}/>}
    {quickAdd!==null&&<QuickAddPartyModal defaultName={quickAdd} type={isSupplier?'supplier':isTransporter?'transporter':'customer'} onClose={()=>setQuickAdd(null)}
      onAdded={p=>{setPId(p.id);setPq('');setQuickAdd(null);foc('acct-amount');}}/>}
    {showReturn&&<SalesReturnModal initialParty={selParty} initialDate={date} onClose={()=>setShowReturn(false)} onSaved={()=>{setShowReturn(false);forceUpdate(n=>n+1);}}/>}
    {editReturnSale&&<EditReturnModal sale={editReturnSale} onClose={()=>setEditReturnSale(null)} onSaved={()=>{setEditReturnSale(null);forceUpdate(n=>n+1);}}/>}

    <SectionHead title="Cash & Accounts" sub={`Net: ${fmt(totalBal)}`}/>

    {/* Account balance cards */}
    <div style={{display:'grid',gridTemplateColumns:mobile?'repeat(2,1fr)':'repeat(4,1fr)',gap:12,marginBottom:18}}>
      {balances.map(a=>(
        <div key={a.id} onClick={()=>setLedger(a.id)}
          style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'14px 16px',
            borderTop:`3px solid ${a.color}`,cursor:'pointer'}}
          {...(mobile?{}:{onMouseEnter:e=>e.currentTarget.style.boxShadow='0 4px 16px rgba(0,0,0,.1)',onMouseLeave:e=>e.currentTarget.style.boxShadow='none'})}>
          <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:4}}>{a.short}</div>
          <div style={{fontSize:20,fontWeight:800,color:a.bal>=0?a.color:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(a.bal)}</div>
          <div style={{fontSize:10,color:C.muted,marginTop:2}}>OB {fmt(a.ob)}</div>
        </div>
      ))}
    </div>

    {/* ── INLINE ENTRY FORM — category-first: pick what it is, direction is auto-detected ── */}
    {/* No overflow:hidden so party dropdown shows above ledger */}
    <div style={{background:C.surf,borderRadius:10,border:`1px solid ${C.bdr}`,marginBottom:18}}>
      {newCat&&<NewCategoryModal onClose={()=>setNewCat(false)} onCreated={cat=>{setNewCat(false);chooseCategory(cat.id);}}/>}

      {/* Field row — order: Date | Category | Party/Payee (or To Account) | Amount | Account | Notes | Save */}
      <div style={{padding:'16px 16px 16px',
        display:mobile?'flex':'grid',flexDirection:mobile?'column':undefined,
        gridTemplateColumns:mobile?undefined:'120px 1.3fr 1.6fr 120px 1fr 1fr auto',
        gridTemplateRows:mobile?undefined:'auto auto',gap:mobile?14:10,position:'relative'}}>

        {/* 1. DATE */}
        <div style={{gridRow:1,alignSelf:'end'}}>
          <label style={labelStyle}>Date <span style={{color:C.muted,fontWeight:400}}>↑↓</span></label>
          <input id="acct-date" type="date" value={date} onChange={e=>setDate(e.target.value)}
            onKeyDown={dateKey} onFocus={()=>setFocused('date')} onBlur={()=>setFocused('')}
            style={fieldStyle(focusedField==='date')}/>
        </div>

        {/* 2. CATEGORY */}
        <div style={{gridRow:1,alignSelf:'end'}}>
          <label style={labelStyle}>Category</label>
          <select id="acct-category" value={categoryId} onChange={e=>chooseCategory(e.target.value)}
            onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();foc('acct-party');}else categoryTypeahead(e);}}
            onFocus={()=>setFocused('category')} onBlur={()=>setFocused('')}
            style={fieldStyle(focusedField==='category',{color:C.txt})}>
            {allCats.map(cx=><option key={cx.id} value={cx.id}>{cx.icon} {cx.label}</option>)}
            <option value="__new__">+ New Category…</option>
          </select>
        </div>

        {/* Direction toggle — sits a row lower, under Category, so it never disturbs the main row's alignment */}
        <div style={{gridRow:2,gridColumn:2,alignSelf:'start'}}>
          {!isTransfer&&<div style={{display:'flex',gap:4}}>
            {['in','out'].map(d=><button key={d} onClick={()=>setDir(d)} type="button" style={{
              flex:1,padding:'3px 6px',borderRadius:5,border:`1px solid ${dir===d?(d==='in'?C.green:C.red):C.bdr}`,
              background:dir===d?(d==='in'?C.greenSoft:C.redSoft):'transparent',
              color:dir===d?(d==='in'?C.green:C.red):C.muted,fontSize:10,fontWeight:700,cursor:'pointer',fontFamily:'inherit'}}>
              {d==='in'?'↑ In':'↓ Out'}
            </button>)}
          </div>}
          {!isTransfer&&dir!==catMeta.dir&&<div style={{fontSize:9,color:C.amber,fontWeight:700,marginTop:4}}>⚠ Unusual for {catMeta.label} — normally {catMeta.dir==='in'?'In':'Out'}</div>}
          {isTransfer&&<div style={{fontSize:9,color:C.muted}}>Contra — moves between your own accounts</div>}
        </div>

        {/* 3. PARTY / PAYEE — or To-Account for a transfer */}
        {isTransfer
          ?<div style={{gridRow:1,alignSelf:'end'}}>
            <label style={labelStyle}>To Account</label>
            <select value={toAccount} onChange={e=>setToAcc(e.target.value)}
              onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();foc('acct-amount');}}}
              style={fieldStyle(false,{color:C.txt})}>
              {ACCOUNTS.filter(a=>a.id!==account).map(a=><option key={a.id} value={a.id}>{a.icon} {a.short}</option>)}
            </select>
          </div>
          :<div style={{gridRow:1,alignSelf:'end',position:'relative'}}>
            <label style={labelStyle}>
              {isSupplier?'Supplier':isTransporter?'Transporter':isParty?'Party':'Payee / Description'}
              {!selParty&&<span style={{color:C.muted,fontWeight:400,marginLeft:5,fontSize:9}}>
                {isSupplier?'e.g. OCB Global':isTransporter?'e.g. Mahamaya Transport':isParty?'e.g. Ravi Traders':'e.g. Pankaj Logistics'}
              </span>}
            </label>
            {selParty&&isParty
              ?<div style={{padding:'7px 11px',borderRadius:7,border:`1.5px solid ${C.green}`,background:C.greenSoft,display:'flex',alignItems:'center',gap:8,boxSizing:'border-box'}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:13,color:C.txt,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{selParty.name}</div>
                  <div style={{fontSize:10,color:C.sub,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
                    {selParty.city}{partyOD!==0?<span style={{color:partyOD>0?C.red:C.green,marginLeft:6}}>{partyOD>0?'OD '+fmt(partyOD):'Credit '+fmt(-partyOD)}</span>:''}
                  </div>
                </div>
                <button onClick={()=>{setPId('');setPq('');setPIdx(0);setSelInvIds([]);setDiscount({});foc('acct-party');}}
                  style={{background:'none',border:'none',cursor:'pointer',color:C.muted,fontSize:18,padding:0,flexShrink:0,lineHeight:1}}>×</button>
              </div>
              :<input id="acct-party" autoComplete="off" autoCorrect="off"
                value={isParty?pq:payee}
                onChange={e=>{ isParty?(setPq(e.target.value),setPIdx(0)):(setPayee(e.target.value),setPIdx(0)); }}
                onFocus={()=>setFocused('party')} onBlur={()=>setTimeout(()=>setFocused(''),150)}
                onKeyDown={e=>{
                  if(isParty){
                    if(e.key==='ArrowDown'){e.preventDefault();setPIdx(i=>Math.min(i+1,psugg.length-1));}
                    if(e.key==='ArrowUp')  {e.preventDefault();setPIdx(i=>Math.max(i-1,0));}
                    if(e.key==='Enter'&&psugg.length){e.preventDefault();const p=psugg[pIdx]||psugg[0];setPId(p.id);setPq('');foc('acct-amount');}
                    else if(e.key==='Enter'&&showAdd){e.preventDefault();setQuickAdd(pq);}
                    else if(e.key==='Enter'){e.preventDefault();foc('acct-amount');}
                  } else {
                    if(e.key==='ArrowDown'){e.preventDefault();setPIdx(i=>Math.min(i+1,payeeSugg.length-1));}
                    if(e.key==='ArrowUp')  {e.preventDefault();setPIdx(i=>Math.max(i-1,0));}
                    if(e.key==='Enter'&&payeeSugg.length&&!payeeSugg.some(p=>p.name.toLowerCase()===payee.trim().toLowerCase())){e.preventDefault();setPayee((payeeSugg[pIdx]||payeeSugg[0]).name);foc('acct-amount');}
                    else if(e.key==='Enter'){e.preventDefault();foc('acct-amount');}
                  }
                }}
                placeholder={isSupplier?'Type supplier name…':isTransporter?'Type transporter name…':isParty?'Type party name…':'Durgesh / OCB Global'}
                style={fieldStyle(focusedField==='party')}/>}
            {/* Dropdown — high z-index so it floats above ledger */}
            {isParty&&!selParty&&(psugg.length>0||showAdd)&&(
              <div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,
                border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:2000,
                boxShadow:'0 12px 32px rgba(0,0,0,.18)',overflow:'hidden',maxHeight:220,overflowY:'auto'}}>
                {psugg.map((p,i)=>(
                  <div key={p.id} onMouseDown={()=>{setPId(p.id);setPq('');foc('acct-amount');}} onMouseEnter={()=>setPIdx(i)}
                    style={{padding:'9px 14px',cursor:'pointer',background:i===pIdx?C.blueSoft:'transparent',borderBottom:`1px solid ${C.bdr}`,fontSize:13}}>
                    <span style={{fontWeight:600,color:C.txt}}>{p.name}</span>
                    <span style={{fontSize:11,color:C.sub,marginLeft:8}}>{p.city}</span>
                    {(()=>{const _od=p.type==='supplier'||p.type==='transporter'?getSupplierBal(p.id):getCustomerBal(p.id);return _od!==0&&<span style={{fontSize:11,color:_od>0?C.red:C.green,marginLeft:8,fontWeight:600}}>{_od>0?'OD '+fmt(_od):'Credit '+fmt(-_od)}</span>;})()}
                  </div>
                ))}
                {showAdd&&<div onMouseDown={()=>setQuickAdd(pq)}
                  style={{padding:'9px 14px',cursor:'pointer',color:C.blue,fontWeight:600,fontSize:13,background:C.blueSoft}}>
                  + Add "{pq}" as new {isSupplier?'supplier':isTransporter?'transporter':'party'}
                </div>}
              </div>
            )}
            {/* Payee pick-list — non-party categories (Salary, Ghar Kharch, etc.):
                names reused from this category's own past entries */}
            {!isParty&&focusedField==='party'&&payeeSugg.length>0&&(
              <div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,
                border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:2000,
                boxShadow:'0 12px 32px rgba(0,0,0,.18)',overflow:'hidden',maxHeight:220,overflowY:'auto'}}>
                {payeeSugg.map((p,i)=>(
                  <div key={p.name} onMouseDown={()=>{setPayee(p.name);foc('acct-amount');}} onMouseEnter={()=>setPIdx(i)}
                    style={{padding:'9px 14px',cursor:'pointer',background:i===pIdx?C.blueSoft:'transparent',borderBottom:`1px solid ${C.bdr}`,fontSize:13,display:'flex',justifyContent:'space-between'}}>
                    <span style={{fontWeight:600,color:C.txt}}>{p.name}</span>
                    <span style={{fontSize:11,color:C.muted}}>used {p.count}×</span>
                  </div>
                ))}
              </div>
            )}
          </div>}

        {/* 4. AMOUNT */}
        <div style={{gridRow:1,alignSelf:'end'}}>
          <label style={labelStyle}>Amount ₹ <span style={{color:C.muted,fontWeight:400}}>e.g. 5000</span></label>
          <div style={{position:'relative'}}>
            <span style={{position:'absolute',left:9,top:'50%',transform:'translateY(-50%)',fontWeight:800,color:C.sub,fontSize:13,pointerEvents:'none'}}>₹</span>
            <input id="acct-amount" type="number" value={amount} onChange={e=>setAmount(e.target.value)}
              onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();foc('acct-account');}}}
              onFocus={()=>setFocused('amount')} onBlur={()=>setFocused('')}
              placeholder="0"
              style={{...fieldStyle(focusedField==='amount'),paddingLeft:22,fontSize:15,fontWeight:700}}/>
          </div>
        </div>

        {/* 5. ACCOUNT */}
        <div style={{gridRow:1,alignSelf:'end'}}>
          <label style={labelStyle}>{isTransfer?'From Account':'Account'}</label>
          <select id="acct-account" value={account} onChange={e=>{
              const newAcc=e.target.value;
              setAccount(newAcc);
              // "To Account" only ever renders accounts other than the current
              // "From Account", so if From is changed to match whatever To
              // already was, To's rendered value silently falls out of sync
              // with its own state (the option it's bound to disappears from
              // the list) — Save then compares the two *state* values, sees
              // them still equal, and refuses with "From and To must differ"
              // even though the dropdowns looked fine. Nudge To off of it here.
              if(isTransfer&&newAcc===toAccount){
                const fallback=ACCOUNTS.find(a=>a.id!==newAcc);
                if(fallback)setToAcc(fallback.id);
              }
            }}
            onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();foc('acct-notes');}}}
            onFocus={()=>setFocused('account')} onBlur={()=>setFocused('')}
            style={fieldStyle(focusedField==='account',{color:C.txt})}>
            {ACCOUNTS.map(a=><option key={a.id} value={a.id}>{a.icon} {a.short} {fmt(getAccBal(a.id))}</option>)}
          </select>
        </div>

        {/* 6. NOTES */}
        <div style={{gridRow:1,alignSelf:'end'}}>
          <label style={labelStyle}>Notes <span style={{color:C.muted,fontWeight:400}}>optional</span></label>
          <input id="acct-notes" value={notes} onChange={e=>setNotes(e.target.value)}
            onFocus={()=>setFocused('notes')} onBlur={()=>setFocused('')}
            onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();doSave();}}}
            placeholder="Inv # / reference"
            style={fieldStyle(focusedField==='notes')}/>
        </div>

        {/* 7. SAVE */}
        <div style={{gridRow:1,alignSelf:'end',paddingTop:mobile?0:18}}>
          <button onClick={doSave} disabled={!canSave}
            style={{padding:'9px 18px',borderRadius:8,border:'none',fontFamily:'inherit',fontWeight:700,fontSize:14,
              background:!canSave?'#E5E7EB':dir==='out'?C.red:dir==='contra'?C.purple:C.green,
              color:!canSave?C.muted:'#fff',
              cursor:!canSave?'not-allowed':'pointer',
              whiteSpace:'nowrap',width:mobile?'100%':undefined,minWidth:80}}>
            Save ↵
          </button>
          {amount>0&&<div style={{fontSize:10,fontWeight:600,marginTop:4,textAlign:'center',
            color:dir==='out'?C.red:dir==='contra'?C.purple:C.green,whiteSpace:'nowrap'}}>
            {fmt(parseFloat(amount)||0)}
          </div>}
        </div>

      </div>
    </div>

    {/* ── INVOICE PICKER — appears between form and ledger when party selected ── */}
    {selParty&&isParty&&(()=>{
      const hasDue=partyBills.length>0||partyOD>0;
      const unpaidBills=partyBills.filter(b=>!b.paid);
      const billNoun=isSupplier?'Bills':isTransporter?'Bills':'Invoices';

      return <div style={{background:'#F0FDF4',borderRadius:10,border:`1.5px solid #86EFAC`,padding:'14px 18px',marginBottom:18}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
          <div style={{fontSize:13,fontWeight:700,color:C.green}}>
            Outstanding {billNoun}
            <span style={{fontSize:11,fontWeight:400,color:C.sub,marginLeft:8}}>FIFO — oldest settled first · select to override{isCustomerParty?' · add discount':''}</span>
          </div>
          <div style={{display:'flex',gap:8}}>
            {selInvIds.length>0&&<button onClick={()=>{
              const selTotal=partyBills.filter(x=>selInvIds.includes(x.key)).reduce((s,x)=>s+(x.remaining||0),0);
              setAmount(String(Math.round(selTotal)));
            }} style={{padding:'4px 12px',borderRadius:6,border:`1px solid ${C.green}`,background:C.surf,color:C.green,fontWeight:700,fontSize:12,cursor:'pointer',fontFamily:'inherit'}}>
              Use selected total
            </button>}
            {isCustomerParty&&<button onClick={()=>setShowReturn(true)} style={{padding:'4px 12px',borderRadius:6,border:`1px solid ${C.amber}`,background:C.surf,color:C.amber,fontWeight:700,fontSize:12,cursor:'pointer',fontFamily:'inherit'}}>
              ↩ Add Return
            </button>}
          </div>
        </div>

        {!hasDue&&<div style={{fontSize:12,color:C.sub,padding:'8px 0'}}>No outstanding dues</div>}
        {hasDue&&partyBills.length===0&&<div style={{fontSize:12,color:C.sub,padding:'8px 0'}}>No {billNoun.toLowerCase()} breakdown — total {isCustomerParty?'OD':'payable'}: <b style={{color:C.red}}>{fmt(partyOD)}</b></div>}

        {unpaidBills.map(b=>{
          const sel=selInvIds.includes(b.key);
          const discKey='disc_'+b.key;
          const repKey='rep_'+b.key;
          const discAmt=isCustomerParty?(parseFloat(discount[discKey]||0)||0):0;
          const repAmt=isCustomerParty?(parseFloat(repair[repKey]||0)||0):0;
          const netOutstanding=Math.max(0,b.remaining-discAmt-repAmt);
          return <div key={b.key} style={{
            display:mobile?'flex':'grid',flexDirection:mobile?'column':undefined,
            gridTemplateColumns:mobile?undefined:'auto 1fr auto auto',gap:mobile?8:12,alignItems:mobile?undefined:'center',
            padding:'10px 12px',borderRadius:8,marginBottom:6,
            border:`1.5px solid ${sel?C.blue:C.bdr}`,
            background:sel?C.blueSoft:C.surf,
            transition:'all .12s'}}>
            {/* Checkbox + date — display:contents on desktop so checkbox and bill-info
                stay as two separate grid items (unchanged layout); a real flex row on mobile */}
            <div style={mobile?{display:'flex',alignItems:'center',gap:10}:{display:'contents'}}>
            <input type="checkbox" checked={sel} onChange={e=>{
              setSelInvIds(ids=>e.target.checked?[...ids,b.key]:ids.filter(k=>k!==b.key));
            }} style={{cursor:'pointer',width:16,height:16,flexShrink:0}}/>

            {/* Bill info */}
            <div>
              <div style={{display:'flex',alignItems:'center',gap:8}}>
                <span style={{fontWeight:700,color:C.txt,fontSize:13}}>{fmtD(b.date)}</span>
                {b.overdue>0
                  ?<span style={{fontSize:10,fontWeight:700,padding:'2px 7px',borderRadius:20,background:C.red,color:'#fff'}}>{b.overdue}d overdue</span>
                  :<span style={{fontSize:10,fontWeight:700,padding:'2px 7px',borderRadius:20,background:C.greenSoft,color:C.green}}>In credit</span>}
                {b.isNew&&<span style={{fontSize:10,fontWeight:700,color:C.blue}}>NEW</span>}
              </div>
              <div style={{fontSize:11,color:C.sub,marginTop:1}}>{b.items||'—'}</div>
            </div>
            </div>

            {/* Per-invoice discount + repair expenditure — customer invoices only;
                suppliers/transporters use the single aggregate write-off box below instead. */}
            {isCustomerParty
              ?<div style={{display:'flex',flexDirection:'column',gap:5}}>
                <div style={{display:'flex',alignItems:'center',gap:6}}>
                  <span style={{fontSize:11,color:C.sub,whiteSpace:'nowrap'}}>Discount ₹</span>
                  <input type="number" placeholder="0"
                    value={discount[discKey]||''}
                    onChange={e=>setDiscount(d=>({...d,[discKey]:e.target.value}))}
                    style={{width:80,padding:'5px 8px',borderRadius:6,border:`1px solid ${C.bdr}`,fontSize:12,outline:'none',fontFamily:'inherit',textAlign:'right'}}
                    onFocus={e=>e.target.style.borderColor=C.amber} onBlur={e=>e.target.style.borderColor=C.bdr}/>
                </div>
                <div style={{display:'flex',alignItems:'center',gap:6}}>
                  <span style={{fontSize:11,color:C.sub,whiteSpace:'nowrap'}}>Repair ₹</span>
                  <input type="number" placeholder="0"
                    value={repair[repKey]||''}
                    onChange={e=>setRepair(d=>({...d,[repKey]:e.target.value}))}
                    style={{width:80,padding:'5px 8px',borderRadius:6,border:`1px solid ${C.bdr}`,fontSize:12,outline:'none',fontFamily:'inherit',textAlign:'right'}}
                    onFocus={e=>e.target.style.borderColor=C.purple} onBlur={e=>e.target.style.borderColor=C.bdr}/>
                </div>
                {repAmt>0&&<div style={{display:'flex',alignItems:'center',gap:6}}>
                  <span style={{fontSize:10,color:C.muted,whiteSpace:'nowrap'}}>↳ Expense ₹</span>
                  <input type="number" placeholder={String(repAmt)}
                    value={repairExpOv[repKey]||''}
                    onChange={e=>setRepairExpOv(d=>({...d,[repKey]:e.target.value}))}
                    style={{width:70,padding:'4px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:11,outline:'none',fontFamily:'inherit',textAlign:'right',color:C.muted}}
                    onFocus={e=>e.target.style.borderColor=C.purple} onBlur={e=>e.target.style.borderColor=C.bdr}/>
                </div>}
              </div>
              :<div/>}

            {/* Net outstanding */}
            <div style={{textAlign:mobile?'left':'right',minWidth:80}}>
              {b.remaining!==b.orig&&<div style={{fontSize:10,color:C.sub,textDecoration:'line-through'}}>{fmt(b.orig)}</div>}
              {(discAmt>0||repAmt>0)&&<div style={{fontSize:10,color:C.amber,textDecoration:'line-through'}}>{fmt(b.remaining)}</div>}
              <div style={{fontWeight:800,color:netOutstanding>0?C.red:C.green,fontSize:14,fontVariantNumeric:'tabular-nums'}}>{fmt(netOutstanding)}</div>
            </div>
          </div>;
        })}

        {/* Summary row */}
        {(partyBills.length>0||partyOD>0)&&<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',
          borderTop:`1px dashed #86EFAC`,marginTop:8,paddingTop:8,fontSize:12}}>
          <div style={{color:C.sub}}>
            {selInvIds.length>0&&<span>Selected: <b style={{color:C.txt}}>{selInvIds.length} bill{selInvIds.length>1?'s':''}</b></span>}
            {isCustomerParty&&Object.values(discount||{}).some(v=>parseFloat(v)>0)&&
              <span style={{marginLeft:12,color:C.amber}}>Total discount: <b>{fmt(Object.values(discount||{}).reduce((s,v)=>s+(parseFloat(v)||0),0))}</b></span>}
            {isCustomerParty&&repairTotal>0&&
              <span style={{marginLeft:12,color:C.purple}}>Total repair: <b>{fmt(repairTotal)}</b></span>}
          </div>
          <div style={{color:(partyOD-(parseFloat(amount)||0)-(isCustomerParty?Object.values(discount||{}).reduce((s,v)=>s+(parseFloat(v)||0),0):0)-(isCustomerParty?repairTotal:0))>0?C.red:C.green,fontWeight:800,fontSize:16}}>
            Net {isCustomerParty?'outstanding':'payable'}: {fmt(partyOD - (parseFloat(amount)||0) - (isCustomerParty?Object.values(discount||{}).reduce((s,v)=>s+(parseFloat(v)||0),0):0) - (isCustomerParty?repairTotal:0))}
          </div>
        </div>}

        {partyReturns.length>0&&<div style={{borderTop:`1px dashed #86EFAC`,marginTop:10,paddingTop:10}}>
          <div style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:6}}>Recent Returns</div>
          {partyReturns.map(r=><div key={r.id} style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'6px 10px',borderRadius:6,marginBottom:4,background:C.redSoft}}>
            <div>
              <span style={{fontWeight:600,color:C.txt,fontSize:12}}>{fmtD(r.date)}</span>
              <span style={{fontSize:11,color:C.sub,marginLeft:8}}>{(r.items||[]).slice(0,2).map(x=>x.customerName||x.itemName).join(', ')}{(r.items||[]).length>2?` +${r.items.length-2} more`:''}</span>
            </div>
            <div style={{display:'flex',alignItems:'center',gap:8}}>
              <span style={{fontWeight:700,color:C.red,fontSize:13,fontVariantNumeric:'tabular-nums'}}>−{fmt(Math.abs(r.total))}</span>
              <button onClick={()=>setEditReturnSale(r)} style={{background:'none',border:'none',cursor:'pointer',color:C.muted,fontSize:13,padding:'0 4px'}} title="Edit return">✎</button>
            </div>
          </div>)}
        </div>}
      </div>;
    })()}

    {/* ── PURCHASE WRITE-OFF — vendor payments have no per-invoice breakdown
         (purchase history has no outstanding-per-bill field like customer
         invoices do), so this is a single aggregate discount against the
         running payable rather than a per-invoice picker. ── */}
    {selParty&&isSupplier&&partyOD!==0&&(()=>{
      const suppDisc=parseFloat(supplierDiscount)||0;
      return <div style={{background:'#FFF7ED',borderRadius:10,border:`1.5px solid #FDBA74`,padding:'14px 18px',marginBottom:18}}>
        <div style={{fontSize:13,fontWeight:700,color:C.amberBright,marginBottom:10}}>
          Purchase Discount / Write-off
          <span style={{fontSize:11,fontWeight:400,color:C.sub,marginLeft:8}}>reduces total purchase amount owed to this supplier</span>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:10,flexWrap:'wrap'}}>
          <div style={{display:'flex',alignItems:'center',gap:6}}>
            <span style={{fontSize:12,color:C.sub}}>Discount ₹</span>
            <input type="number" placeholder="0" value={supplierDiscount} onChange={e=>setSupplierDiscount(e.target.value)}
              style={{width:100,padding:'6px 9px',borderRadius:6,border:`1px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit',textAlign:'right'}}
              onFocus={e=>e.target.style.borderColor=C.amber} onBlur={e=>e.target.style.borderColor=C.bdr}/>
          </div>
          <div style={{marginLeft:'auto',fontSize:12,color:partyOD-suppDisc>0?C.red:C.green,fontWeight:700}}>
            Net payable: {fmt(partyOD-suppDisc)}
          </div>
        </div>
      </div>;
    })()}

    {/* ── LEDGER — category-wise entries: pick a category to see its own in/out ledger ── */}
    {editTxn&&<EditTxnModal txn={editTxn} onClose={()=>setEditTxn(null)} onSaved={()=>forceUpdate(n=>n+1)}/>}
    {catLedger&&<CategoryLedger catId={catLedger} onClose={()=>setCatLedger(null)}/>}
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',maxHeight:480,overflowY:'auto'}}>
      <div style={{padding:'11px 16px',borderBottom:`1px solid ${C.bdr}`,display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
        <span style={{fontSize:14,fontWeight:700,color:C.txt,marginRight:4}}>Transactions · {filtered.length}</span>
        <input value={srch} onChange={e=>{setSrch(e.target.value);setPage(0);}} placeholder="Search…"
          style={{flex:1,minWidth:140,padding:'6px 10px',borderRadius:6,border:`1px solid ${C.bdr}`,fontSize:12,outline:'none',fontFamily:'inherit'}}/>
        <select value={catFilter} onChange={e=>{setCatFilter(e.target.value);setPage(0);}} onKeyDown={catFilterTypeahead}
          style={{padding:'6px 10px',borderRadius:6,border:`1px solid ${catFilter?C.blue:C.bdr}`,color:catFilter?C.blue:C.sub,fontSize:12,fontWeight:600,fontFamily:'inherit',background:catFilter?C.blueSoft:C.surf}}>
          <option value="">All categories</option>
          {allCats.map(cx=><option key={cx.id} value={cx.id}>{cx.icon} {cx.label}</option>)}
        </select>
      </div>
      {catTotals&&<div style={{display:'flex',gap:16,padding:'8px 16px',borderBottom:`1px solid ${C.bdr}`,background:'#F9FAFB',fontSize:12,alignItems:'center'}}>
        <span style={{color:C.sub}}>{getCatLabel(catFilter)} (this app) —</span>
        <span style={{color:C.green,fontWeight:700}}>In {fmt(catTotals.inSum)}</span>
        <span style={{color:C.red,fontWeight:700}}>Out {fmt(catTotals.outSum)}</span>
        <span style={{color:catTotals.net>=0?C.txt:C.red,fontWeight:700}}>Net {catTotals.net<0?'−':''}{fmt(Math.abs(catTotals.net))}</span>
        <button onClick={()=>setCatLedger(catFilter)} style={{marginLeft:'auto',padding:'3px 10px',borderRadius:6,border:`1px solid ${C.blue}`,background:C.surf,color:C.blue,fontWeight:700,fontSize:11,cursor:'pointer',fontFamily:'inherit'}}>View Full Ledger</button>
      </div>}
      <CardTable mobile={mobile} rows={pageData} keyFn={(t,i)=>t.id||i} emptyLabel="No transactions yet"
        columns={[
          {key:'date',label:'Date',render:t=><span style={{color:C.sub,whiteSpace:'nowrap'}}>{fmtD(t.date)}</span>},
          {key:'type',label:'Type',render:t=><span style={{fontSize:11,fontWeight:700,padding:'2px 8px',borderRadius:20,background:(TYPE_COLOR[t.type]||C.blue)+'18',color:TYPE_COLOR[t.type]||C.blue}}>{TYPE_LABEL[t.type]||t.type}</span>},
          {key:'account',label:'Account',render:t=><span><span style={{fontSize:12,fontWeight:600,color:ACC_MAP[t.account]?.color||C.sub}}>{ACC_MAP[t.account]?.short}</span>{t.toAccount&&<span style={{fontSize:11,color:C.muted}}> → {ACC_MAP[t.toAccount]?.short}</span>}</span>},
          {key:'desc',label:'Description',render:t=>{
            const desc=t.type==='transfer'?`${ACC_MAP[t.account]?.short} → ${ACC_MAP[t.toAccount]?.short}`
              :t.type==='party_payment'?(t.partyName||t.payee||'—')
              :t.payee||t.partyName||(t.category?getCatLabel(t.category):'—');
            return <div><div style={{fontWeight:500,color:C.txt}}>{desc}</div>{t.notes&&<div style={{fontSize:11,color:C.muted}}>{t.notes}</div>}{t.category&&<div style={{fontSize:10,color:C.muted}}>{getCatLabel(t.category)}</div>}</div>;
          }},
          {key:'amount',label:'Amount',align:'right',render:t=>{
            const isOut=(t.type==='payment')===(t.amount>=0);
            return <span style={{fontWeight:700,fontVariantNumeric:'tabular-nums',color:isOut?C.red:C.green}}>{isOut?'−':'+'}{fmt(Math.abs(t.amount))}</span>;
          }},
          {key:'actions',label:'',cardLabel:'Actions',render:t=>
            <div style={{display:'flex',gap:6,justifyContent:'flex-end'}}>
              <Btn onClick={()=>setEditTxn(t)} variant="ghost" size="sm">Edit</Btn>
              <Btn onClick={()=>deleteTxn(t.id)} variant="ghost" size="sm">Delete</Btn>
            </div>},
        ]}/>
      {totalPages>1&&<div style={{padding:'10px 16px',borderTop:`1px solid ${C.bdr}`,display:'flex',justifyContent:'center',gap:8}}>
        <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
        <span style={{fontSize:12,color:C.sub,alignSelf:'center'}}>Page {page+1} of {totalPages}</span>
        <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
      </div>}
    </div>
  </div>;
}

function AccountLedger({accId,onClose}) {
  const acc=ACC_MAP[accId];
  const [search,setSearch]=useState('');
  const [page,setPage]=useState(0);
  const mobile=useIsMobile();
  const PAGE=20;
  const histEntries=useMemo(()=>ACCOUNT_HISTORY[accId]||[],[accId]);
  const newTxns=useMemo(()=>state.txns.filter(t=>t.account===accId||t.toAccount===accId),[accId,state.txns.length]);

  // acc.ob is the balance as of when this app started — i.e. AFTER all the
  // historical entries below already happened. To show one continuous running
  // balance across historical + new, walk forward from a computed "true"
  // opening balance (ob minus the historical entries' net effect) through
  // both, oldest first, landing back on ob right before the new entries begin.
  const histNet=useMemo(()=>histEntries.reduce((s,e)=>s+e[1],0),[histEntries]);
  const trueOpening=acc.ob-histNet;

  const allRows=useMemo(()=>{
    const events=[
      ...histEntries.map(([d,amt,desc])=>({date:d,delta:amt,desc,isHist:true,sortKey:d})),
      ...newTxns.map(t=>{
        const isIn=t.type==='transfer'?t.toAccount===accId:(t.type==='receipt'||t.type==='party_payment');
        const delta=t.type==='transfer'?(t.account===accId?-t.amount:t.amount):((t.type==='receipt'||t.type==='party_payment')?t.amount:-t.amount);
        const isVendorLike=t.category==='vendor'||t.category==='transporter_payment';
        const desc=t.type==='transfer'?(isIn?`From ${ACC_MAP[t.account]?.short}`:`To ${ACC_MAP[t.toAccount]?.short}`)
          :t.type==='party_payment'?'Party — '+(t.partyName||t.payee)
          :isVendorLike?getCatLabel(t.category)+' — '+(t.partyName||t.payee)
          :t.partyName?'Receipt — '+t.partyName
          :(t.category?getCatLabel(t.category)+' — ':'')+(t.payee||'');
        return {date:t.date,delta,desc,notes:t.notes,isHist:false,sortKey:t.date+'_'+(t.createdAt||'')};
      }),
    ].sort((a,b)=>a.sortKey.localeCompare(b.sortKey));
    let running=trueOpening;
    return events.map(e=>{running+=e.delta;return {...e,balance:running};}).reverse();
  },[histEntries,newTxns,trueOpening,accId]);

  const filtered=useMemo(()=>allRows.filter(r=>!search||r.desc.toLowerCase().includes(search.toLowerCase())||(r.notes||'').toLowerCase().includes(search.toLowerCase())),[allRows,search]);
  const pageRows=filtered.slice(page*PAGE,(page+1)*PAGE);
  const totalPages=Math.ceil(filtered.length/PAGE);
  return <Modal title={acc.label} onClose={onClose} width={640}>
    <div style={{display:'flex',flexDirection:mobile?'column':'row',gap:mobile?14:0,justifyContent:'space-between',alignItems:mobile?'flex-start':'center',marginBottom:16}}>
      <div>
        <div style={{fontSize:11,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>Current Balance</div>
        <div style={{fontSize:28,fontWeight:800,color:getAccBal(accId)>=0?acc.color:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(getAccBal(accId))}</div>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,textAlign:mobile?'left':'right',width:mobile?'100%':undefined}}>
        <div><div style={{fontSize:10,color:C.sub}}>Opening (pre-app)</div><div style={{fontWeight:700,color:C.txt}}>{fmt(trueOpening)}</div></div>
        <div><div style={{fontSize:10,color:C.sub}}>Entries</div><div style={{fontWeight:700,color:C.txt}}>{allRows.length}</div></div>
        <div><div style={{fontSize:10,color:C.sub}}>Total IN</div><div style={{fontWeight:700,color:C.green}}>{fmt(allRows.filter(r=>r.delta>0).reduce((s,r)=>s+r.delta,0))}</div></div>
        <div><div style={{fontSize:10,color:C.sub}}>Total OUT</div><div style={{fontWeight:700,color:C.red}}>{fmt(Math.abs(allRows.filter(r=>r.delta<0).reduce((s,r)=>s+r.delta,0)))}</div></div>
      </div>
    </div>
    <div style={{background:C.surf,borderRadius:7,border:`1px solid ${C.bdr}`,padding:'7px 11px',marginBottom:12}}>
      <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="Search…" style={{width:'100%',border:'none',outline:'none',fontSize:13,color:C.txt,fontFamily:'inherit',background:'transparent'}}/>
    </div>
    <div style={{border:`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden'}}>
      <CardTable mobile={mobile} rows={pageRows} keyFn={(t,i)=>i} emptyLabel="No transactions"
        columns={[
          {key:'date',label:'Date',render:t=><span style={{color:C.sub,whiteSpace:'nowrap'}}>{fmtD(t.date)}</span>},
          {key:'desc',label:'Description',render:t=><div><div style={{fontWeight:500,color:C.txt}}>{t.desc}{t.isHist&&<span style={{fontSize:10,color:C.muted,marginLeft:6}}>pre-app</span>}</div>{t.notes&&<div style={{fontSize:11,color:C.muted}}>{t.notes}</div>}</div>},
          {key:'in',label:'In ↑',align:'right',render:t=><span style={{fontWeight:700,color:C.green,fontVariantNumeric:'tabular-nums'}}>{t.delta>0?fmt(t.delta):'—'}</span>},
          {key:'out',label:'Out ↓',align:'right',render:t=><span style={{fontWeight:700,color:C.red,fontVariantNumeric:'tabular-nums'}}>{t.delta<=0?fmt(Math.abs(t.delta)):'—'}</span>},
          {key:'balance',label:'Balance',align:'right',render:t=><span style={{fontWeight:700,color:t.balance>=0?C.txt:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(t.balance)}</span>},
        ]}/>
    </div>
    {totalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:12}}>
      <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {page+1} of {totalPages} · {filtered.length} entries</span>
      <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}
  </Modal>;
}

function CategoryLedger({catId,onClose}) {
  const label=getCatLabel(catId);
  const catMeta=getAllCats().find(c=>c.id===catId)||{dir:'out'};
  const historical=HISTORICAL_EXPENSES[catId]||0;
  const [search,setSearch]=useState('');
  const [page,setPage]=useState(0);
  const [groupView,setGroupView]=useState(false);
  const [expandedName,setExpandedName]=useState(null);
  const mobile=useIsMobile();
  const subGroups=useMemo(()=>getCategorySubLedger(catId),[catId,state.txns.length]);
  const PAGE=20;
  const matchesCat=t=>{
    if(catId==='party_receipt') return t.type==='party_payment';
    if(catId==='other_receipt') return t.type==='receipt';
    if(catId==='transfer') return t.type==='transfer';
    return t.category===catId;
  };
  const txns=useMemo(()=>state.txns.filter(t=>matchesCat(t)&&(!search||(t.payee||'').toLowerCase().includes(search.toLowerCase())||(t.partyName||'').toLowerCase().includes(search.toLowerCase())||(t.notes||'').toLowerCase().includes(search.toLowerCase())))
    .sort((a,b)=>b.date.localeCompare(a.date)||((b.createdAt||'').localeCompare(a.createdAt||''))),[catId,search,state.txns.length]);

  // Two different numbers per entry:
  //  - isOut: the literal cash direction (money actually paid out vs received) — drives
  //    the In/Out columns and the Total In/Out summary, same convention as everywhere else.
  //  - catDelta: contribution to the category's own cumulative total (e.g. total spent, for
  //    an expense category). An entry matching the category's home direction ADDS to that
  //    running total; a reversed/overridden entry (e.g. a refund) SUBTRACTS — the OPPOSITE
  //    convention from a cash-account balance, where money in always adds.
  let running=historical; const rows=[];
  for(const t of [...txns].reverse()){
    const isOut=(t.type==='payment')===(t.amount>=0);
    const catDelta=(isOut===(catMeta.dir==='out'))?Math.abs(t.amount):-Math.abs(t.amount);
    running+=catDelta; rows.unshift({...t,isOut,catDelta,balance:running});
  }
  const totalIn=rows.filter(r=>!r.isOut).reduce((s,r)=>s+Math.abs(r.amount),0);
  const totalOut=rows.filter(r=>r.isOut).reduce((s,r)=>s+Math.abs(r.amount),0);
  const pageRows=rows.slice(page*PAGE,(page+1)*PAGE);
  const totalPages=Math.ceil(rows.length/PAGE);

  return <Modal title={label+' — Ledger'} onClose={onClose} width={640}>
    <div style={{display:'flex',flexDirection:mobile?'column':'row',gap:mobile?14:0,justifyContent:'space-between',alignItems:mobile?'flex-start':'center',marginBottom:16}}>
      <div>
        <div style={{fontSize:11,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>{catMeta.dir==='out'?'Total Spent':'Total Collected'} (historical + this app)</div>
        <div style={{fontSize:28,fontWeight:800,color:running>=0?C.txt:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(running)}</div>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,textAlign:mobile?'left':'right',width:mobile?'100%':undefined}}>
        <div><div style={{fontSize:10,color:C.sub}}>Historical</div><div style={{fontWeight:700,color:C.txt}}>{fmt(historical)}</div></div>
        <div><div style={{fontSize:10,color:C.sub}}>Entries (this app)</div><div style={{fontWeight:700,color:C.txt}}>{txns.length}</div></div>
        <div><div style={{fontSize:10,color:C.sub}}>Total In</div><div style={{fontWeight:700,color:C.green}}>{fmt(totalIn)}</div></div>
        <div><div style={{fontSize:10,color:C.sub}}>Total Out</div><div style={{fontWeight:700,color:C.red}}>{fmt(totalOut)}</div></div>
      </div>
    </div>
    <div style={{display:'flex',gap:8,marginBottom:12}}>
      <Chip v="Timeline" active={!groupView} onClick={()=>setGroupView(false)}/>
      <Chip v="By Person / Head" active={groupView} onClick={()=>setGroupView(true)}/>
    </div>
    {groupView?(
      subGroups.length===0
        ?<div style={{padding:'30px',textAlign:'center',color:C.muted,border:`1px solid ${C.bdr}`,borderRadius:8}}>No named breakdown available for this category yet — entries will group here once they share a payee name.</div>
        :<div style={{border:`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden'}}>
          <CardTable mobile={mobile} rows={subGroups} keyFn={g=>g.name} emptyLabel="No named breakdown"
            onRowClick={g=>setExpandedName(expandedName===g.name?null:g.name)}
            renderSubRow={g=>expandedName===g.name?
              <div style={{padding:mobile?'8px 14px':'8px 12px 8px 30px',background:'#F9FAFB'}}>
                {[...g.histEntries.map(e=>({date:e[0],amount:e[1],refund:e[1]<0})),...g.newEntries.map(t=>({date:t.date,amount:Math.abs(t.amount),refund:false}))]
                  .sort((a,b)=>b.date.localeCompare(a.date)).map((e,j)=>(
                  <div key={j} style={{display:'flex',justifyContent:'space-between',padding:'3px 0',fontSize:12,borderBottom:j<g.histEntries.length+g.newEntries.length-1?`1px solid ${C.bdr}`:'none'}}>
                    <span style={{color:C.sub}}>{fmtD(e.date)}{e.refund&&<span style={{color:C.amber,marginLeft:6}}>refund</span>}</span>
                    <span style={{fontWeight:600,color:e.refund?C.red:C.txt}}>{e.refund?'−':''}{fmt(Math.abs(e.amount))}</span>
                  </div>
                ))}
              </div>:null}
            columns={[
              {key:'name',label:'Name',render:g=><span style={{fontWeight:600,color:C.blue}}>{expandedName===g.name?'▾':'▸'} {g.name}</span>},
              {key:'hist',label:'Historical',align:'right',render:g=><span style={{color:g.histTotal<0?C.red:C.sub}}>{g.histTotal<0?'−':''}{fmt(Math.abs(g.histTotal))}</span>},
              {key:'new',label:'This App',align:'right',render:g=><span style={{color:g.newTotal<0?C.red:C.sub}}>{g.newTotal<0?'−':''}{fmt(Math.abs(g.newTotal))}</span>},
              {key:'total',label:'Total',align:'right',render:g=><span style={{fontWeight:700,color:g.total<0?C.red:C.txt}}>{g.total<0?'−':''}{fmt(Math.abs(g.total))}</span>},
            ]}/>
        </div>
    ):<>
    <div style={{background:C.surf,borderRadius:7,border:`1px solid ${C.bdr}`,padding:'7px 11px',marginBottom:12}}>
      <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="Search…" style={{width:'100%',border:'none',outline:'none',fontSize:13,color:C.txt,fontFamily:'inherit',background:'transparent'}}/>
    </div>
    <div style={{border:`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden'}}>
      <CardTable mobile={mobile} rows={pageRows} keyFn={(t,i)=>i} emptyLabel="No entries recorded in this app yet"
        footer={<div style={{display:'flex',justifyContent:'space-between',padding:'9px 12px',background:'#F6F8FA',fontWeight:700,color:C.txt}}>
          <span>Historical (pre-app)</span><span style={{fontWeight:800}}>{fmt(historical)}</span>
        </div>}
        columns={[
          {key:'date',label:'Date',render:t=><span style={{color:C.sub,whiteSpace:'nowrap'}}>{fmtD(t.date)}</span>},
          {key:'desc',label:'Description',render:t=>{
            const desc=t.partyName||t.payee||'—';
            return <div><div style={{fontWeight:500,color:C.txt}}>{desc}</div>{t.notes&&<div style={{fontSize:11,color:C.muted}}>{t.notes}</div>}</div>;
          }},
          {key:'in',label:'In ↑',align:'right',render:t=><span style={{fontWeight:700,color:C.green,fontVariantNumeric:'tabular-nums'}}>{!t.isOut?fmt(Math.abs(t.amount)):'—'}</span>},
          {key:'out',label:'Out ↓',align:'right',render:t=><span style={{fontWeight:700,color:C.red,fontVariantNumeric:'tabular-nums'}}>{t.isOut?fmt(Math.abs(t.amount)):'—'}</span>},
          {key:'balance',label:catMeta.dir==='out'?'Spent so far':'Collected so far',align:'right',render:t=><span style={{fontWeight:700,color:t.balance>=0?C.txt:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(t.balance)}</span>},
        ]}/>
    </div>
    {totalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:12}}>
      <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {page+1} of {totalPages} · {rows.length} entries</span>
      <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}
    </>}
  </Modal>;
}

function Dashboard({navigate}){
  const mobile=useIsMobile();
  const customers=state.parties.filter(p=>p.type==='customer');
  const salesByParty=state.sales.reduce((m,s)=>{m[s.partyId]=(m[s.partyId]||0)+(s.total||0);return m;},{});
  const paymentsByParty=state.payments.reduce((m,p)=>{m[p.partyId]=(m[p.partyId]||0)+(p.amount||0);return m;},{});
  const odParties=customers
    .map(p=>({...p,od:(p.ob||0)+(salesByParty[p.id]||0)-(paymentsByParty[p.id]||0)}))
    .filter(p=>p.od>0).sort((a,b)=>b.od-a.od);
  const totalOD=odParties.reduce((s,p)=>s+p.od,0);
  const outOfStock=state.items.filter(it=>it.active!==false&&(it.stock||0)<=0);
  const monthPrefix=today().slice(0,7);
  const monthSales=state.sales.filter(s=>s.date&&s.date.startsWith(monthPrefix));
  const monthAmt=monthSales.reduce((s,t)=>s+t.total,0);
  // Two-bucket breakup requested for the Sales This Month card — Lighter is
  // its own high-volume category, everything else (Infrared, Induction, Gas
  // Stove, etc.) rolls up into "Appliances".
  const monthCatSplit=monthSales.reduce((acc,s)=>{
    for(const l of (s.items||[])){
      const it=state.items.find(x=>x.id===l.itemId);
      const amt=l.amount??(l.qty*l.rate);
      if(it?.cat==='Lighter')acc.lighter+=amt;else acc.appliances+=amt;
    }
    return acc;
  },{lighter:0,appliances:0});
  const t=today();
  const[_fvD,_setFvD]=useState(0); const forceUpdate=()=>_setFvD(v=>v+1);
  const[ledgerParty,setLedgerParty]=useState(null);
  const followupsDue=state.followups.filter(f=>f.status==='open'&&f.followUpDate&&f.followUpDate<=t).sort((a,b)=>a.followUpDate.localeCompare(b.followUpDate));
  const followupsUpcoming=state.followups.filter(f=>f.status==='open'&&f.followUpDate&&f.followUpDate>t).sort((a,b)=>a.followUpDate.localeCompare(b.followUpDate)).slice(0,5);
  const markDone=id=>{state.followups=state.followups.map(f=>f.id===id?{...f,status:'done'}:f);save.followups();forceUpdate(n=>n+1);};

  const followupsBlock=<div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
    <div style={{padding:'12px 16px',borderBottom:`1px solid ${C.bdr}`,background:'#FFF8C5',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <h3 style={{margin:0,fontSize:14,fontWeight:700,color:C.amberBright}}>📋 Follow-ups · {followupsDue.length} due</h3>
      {followupsUpcoming.length>0&&<span style={{fontSize:11,color:C.sub}}>{followupsUpcoming.length} upcoming</span>}
    </div>
    {[...followupsDue,...followupsUpcoming].length===0
      ?<div style={{padding:'28px',textAlign:'center',color:C.muted,fontSize:13}}>No follow-ups pending ✓</div>
      :[...followupsDue,...followupsUpcoming].slice(0,8).map((f,i)=>{
        const isDue=f.followUpDate<=t;
        return <div key={i} style={{display:'flex',alignItems:'flex-start',gap:10,padding:'10px 16px',borderBottom:`1px solid ${C.bdr}`,background:isDue?'#FFFFF0':'transparent'}}>
          <div style={{flex:1,minWidth:0}}>
            <div onClick={()=>{const p=state.parties.find(pp=>pp.id===f.partyId);if(p)setLedgerParty(p);}}
              style={{fontWeight:600,color:C.blue,fontSize:13,cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted'}}>{f.partyName}</div>
            <div style={{fontSize:11,marginTop:1}}>
              <span style={{color:TAG_COLORS[f.tags[0]]||C.blue,fontWeight:700}}>{f.label}</span>
              <span style={{color:C.sub}}>{' · '}{isDue?'Due today':'Due '+fmtD(f.followUpDate)}</span>
            </div>
            <div style={{fontSize:11,color:C.muted,fontStyle:'italic',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>"{f.response}"</div>
          </div>
          <div style={{display:'flex',gap:5,flexShrink:0}}>
            <button onClick={()=>markDone(f.id)} style={{padding:'3px 8px',borderRadius:5,background:C.greenSoft,color:C.green,border:`1px solid ${C.green}40`,cursor:'pointer',fontSize:11,fontWeight:700,fontFamily:'inherit'}}>Done</button>
            <button onClick={()=>navigate('outstanding')} style={{padding:'3px 8px',borderRadius:5,background:C.blueSoft,color:C.blue,border:`1px solid ${C.blue}40`,cursor:'pointer',fontSize:11,fontWeight:600,fontFamily:'inherit'}}>View</button>
          </div>
        </div>;
      })}
  </div>;
  const topOutstandingBlock=<div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
    <div style={{padding:'12px 16px',borderBottom:`1px solid ${C.bdr}`,display:'flex',justifyContent:'space-between'}}>
      <h3 style={{margin:0,fontSize:14,fontWeight:700,color:C.txt}}>Top Outstanding</h3>
      <button onClick={()=>navigate('outstanding')} style={{background:'none',border:'none',cursor:'pointer',fontSize:12,color:C.blue,fontWeight:600}}>All →</button>
    </div>
    {odParties.length===0
      ?<div style={{padding:'20px',textAlign:'center',color:C.muted,fontSize:13}}>No outstanding dues ✓</div>
      :odParties.slice(0,6).map((p,i)=><div key={i} style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'8px 16px',borderBottom:`1px solid ${C.bdr}`}}>
        <div>
          <div onClick={()=>setLedgerParty(p)} style={{fontWeight:600,color:C.blue,fontSize:12,cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted'}}>{p.name}</div>
          <div style={{fontSize:10,color:C.muted}}>{p.city}·{p.route}</div>
        </div>
        <span style={{fontWeight:700,color:C.red,fontVariantNumeric:'tabular-nums',fontSize:13}}>{fmt(p.od)}</span>
      </div>)}
  </div>;
  const quickActionsBlock=<div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'14px 16px'}}>
    <h3 style={{margin:'0 0 10px',fontSize:14,fontWeight:700,color:C.txt}}>Quick Actions</h3>
    <div style={{display:'grid',gridTemplateColumns:mobile?'repeat(4,1fr)':'1fr 1fr',gap:8}}>
      {[{l:'New Sale',i:'🧾',m:'sale',c:C.blue},{l:'Outstanding',i:'⏰',m:'outstanding',c:C.red},{l:'Parties',i:'🏪',m:'parties',c:C.green},{l:'Inventory',i:'📦',m:'items',c:C.purple}].map(a=>(
        mobile
          ?<button key={a.m} onClick={()=>navigate(a.m)} style={{display:'flex',flexDirection:'column',alignItems:'center',gap:4,padding:'10px 4px',borderRadius:7,border:`1px solid ${a.c}20`,background:a.c+'08',cursor:'pointer',fontSize:10,fontWeight:600,color:a.c,fontFamily:'inherit'}}>
            <span style={{fontSize:19}}>{a.i}</span>{a.l}
          </button>
          :<button key={a.m} onClick={()=>navigate(a.m)} style={{display:'flex',alignItems:'center',gap:8,padding:'9px 12px',borderRadius:7,border:`1px solid ${a.c}20`,background:a.c+'08',cursor:'pointer',fontSize:12,fontWeight:600,color:a.c,fontFamily:'inherit'}}>
            <span style={{fontSize:17}}>{a.i}</span>{a.l}
          </button>
      ))}
    </div>
  </div>;

  return <div>
    {ledgerParty&&<PartyLedger party={ledgerParty} onClose={()=>setLedgerParty(null)}/>}
    <div style={{marginBottom:18}}>
      <h1 style={{margin:0,fontSize:mobile?19:22,fontWeight:800,color:C.txt,letterSpacing:-.5}}>{new Date().getHours()<12?'Good morning':'Good afternoon'}, Pratik 👋</h1>
      <p style={{margin:'3px 0 0',fontSize:13,color:C.sub}}>{new Date().toLocaleDateString('en-IN',{weekday:'long',day:'numeric',month:'long'})}</p>
    </div>
    <div style={{display:'grid',gridTemplateColumns:mobile?'repeat(2,1fr)':'repeat(auto-fit,minmax(175px,1fr))',gap:mobile?10:12,marginBottom:18}}>
      <StatCard label="Total Outstanding" value={fmtL(totalOD)} sub={`${odParties.length} parties`} color={C.red} icon="⏰" onClick={()=>navigate('outstanding')}/>
      <StatCard label="Sales This Month" value={monthSales.length?fmtL(monthAmt):'—'} sub={monthSales.length?<>
        {monthSales.length} invoices
        <div style={{marginTop:2}}>Lighter {fmtL(monthCatSplit.lighter)} · Appliances {fmtL(monthCatSplit.appliances)}</div>
      </>:'No sales'} color={C.green} icon="🧾"/>
      <StatCard label="Follow-ups Due" value={String(followupsDue.length||0)} sub={followupsDue.length?'Action needed':'All clear'} color={followupsDue.length>0?C.amber:C.green} icon="📋" onClick={()=>navigate('outstanding')}/>
      <StatCard label="Parties" value={String(customers.length)} sub={`${state.parties.filter(p=>p.type==='supplier').length} suppliers`} color={C.blue} icon="🏪" onClick={()=>navigate('parties')}/>
      <StatCard label="Out of Stock" value={String(outOfStock.length)} sub={outOfStock.length?'Items need reorder':'All stocked'} color={outOfStock.length>0?C.red:C.green} icon="📦" onClick={()=>navigate('items')}/>
    </div>
    {mobile
      ?<div style={{display:'flex',flexDirection:'column',gap:12}}>
        {quickActionsBlock}
        {followupsBlock}
        {topOutstandingBlock}
      </div>
      :<div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14,alignItems:'start'}}>
        {followupsBlock}
        <div style={{display:'flex',flexDirection:'column',gap:12}}>
          {topOutstandingBlock}
          {quickActionsBlock}
        </div>
      </div>}
  </div>;
}

// ── ITEM MASTERS ──────────────────────────────────────────────────────────────
function CategoryDetail({cat,items,onClose}){
  const mobile=useIsMobile();
  const[search,setSearch]=useState('');
  const[page,setPage]=useState(0);
  const PAGE=20;
  const metrics=useMemo(()=>getAllItemMetrics(),[state.sales.length,state.purchases.length,state.items.length]);
  const expectedMap=useMemo(()=>getExpectedStockMap(),[state.purchases.length]);
  // FIFO-costed stock value — item.purch is a flat rate only refreshed by the
  // offline seed/import script, so it goes stale the moment an in-app
  // purchase lands at a different rate. getFifoStockMap layers real
  // purchase/sale/repair-expense events (dated, itemId-tagged) on top of a
  // reconciled baseline instead — see its definition in store.js.
  const fifoMap=useMemo(()=>getFifoStockMap(),[state.purchases.length,state.sales.length,state.repairExpenses.length,state.items.length]);
  const catMetrics=useMemo(()=>metrics.filter(m=>(m.item.cat||'Other')===cat),[metrics,cat]);
  const b=useMemo(()=>catMetrics.reduce((s,m)=>{
    const stock=m.item.stock||0;
    const fifo=fifoMap[m.item.id];
    s.stock+=stock; s.stockValue+=fifo?fifo.value:(stock*(m.item.purch||0)); s.damaged+=(m.item.damaged||0);
    s.totalSoldQty+=m.totalSoldQty; s.totalRevenue+=m.totalRevenue; s.profit+=m.profit; s.expected+=(expectedMap[m.item.id]||0);
    return s;
  },{stock:0,stockValue:0,damaged:0,totalSoldQty:0,totalRevenue:0,profit:0,expected:0}),[catMetrics,expectedMap,fifoMap]);
  const avgProfitPct=b.totalRevenue>0?Math.round(b.profit/b.totalRevenue*1000)/10:0;
  const itemRows=useMemo(()=>catMetrics.filter(m=>!search||(m.item.internal||m.item.name||'').toLowerCase().includes(search.toLowerCase())).sort((a,b)=>b.totalRevenue-a.totalRevenue),[catMetrics,search]);
  const pageRows=itemRows.slice(page*PAGE,(page+1)*PAGE);
  const totalPages=Math.ceil(itemRows.length/PAGE);
  const detailColumns=[
    {key:'item',label:'Item',render:m=><span style={{fontWeight:600,color:C.txt}}>{m.item.internal||m.item.name}</span>},
    {key:'stock',label:'Stock',align:'right',render:m=><span style={{color:C.sub}}>{m.item.stock||0}</span>},
    {key:'stockValue',label:'Stock Value',align:'right',render:m=>{
      const fifo=fifoMap[m.item.id];
      const v=fifo?fifo.value:((m.item.stock||0)*(m.item.purch||0));
      return <span style={{fontWeight:600,color:C.txt}}>{fmt(v)}</span>;
    }},
    {key:'expected',label:'Expected',align:'right',render:m=>{
      const expected=expectedMap[m.item.id]||0;
      return <span style={{color:expected>0?C.amber:C.muted}}>{expected>0?'+'+expected:'—'}</span>;
    }},
    {key:'sold',label:'Sold',align:'right',render:m=><span style={{color:C.sub}}>{m.totalSoldQty}</span>},
    {key:'revenue',label:'Revenue',align:'right',render:m=><span style={{fontWeight:600,color:C.txt}}>{fmt(m.totalRevenue)}</span>},
    {key:'profit',label:'Profit',align:'right',render:m=><span style={{fontWeight:600,color:m.profit>=0?C.blue:C.red}}>{fmt(m.profit)}</span>},
    {key:'profitPct',label:'Profit %',align:'right',render:m=>{
      const pct=m.totalRevenue>0?Math.round(m.profit/m.totalRevenue*1000)/10:0;
      return <Badge v={pct+'%'} color={pct>25?'green':pct>15?'amber':'red'}/>;
    }},
    {key:'damaged',label:'Damaged',align:'right',render:m=><span style={{color:(m.item.damaged||0)>0?C.amber:C.sub}}>{m.item.damaged||0}</span>},
  ];
  return <Modal title={cat+' — Category Detail'} onClose={onClose} width={720}>
    <div style={{display:'grid',gridTemplateColumns:mobile?'repeat(2,1fr)':'repeat(4,1fr)',gap:8,marginBottom:16}}>
      {[['Items',catMetrics.length],['Stock (pcs)',b.stock],['Stock Value',fmt(b.stockValue)],['Damaged Stock',b.damaged+' pcs'],
        ['Total Sold (pcs)',b.totalSoldQty],['Total Revenue',fmt(b.totalRevenue)],['Profit Earned',fmt(b.profit)],['Avg Profit %',avgProfitPct+'%'],['Expected (Incoming)',b.expected+' pcs']]
        .map(([label,value],i)=><div key={i} style={{background:'#F9FAFB',borderRadius:7,padding:'8px 10px',border:`1px solid ${C.bdr}`}}>
          <div style={{fontSize:9,color:C.sub,textTransform:'uppercase',letterSpacing:.4}}>{label}</div>
          <div style={{fontSize:14,fontWeight:800,color:C.txt}}>{value}</div>
        </div>)}
    </div>
    <div style={{background:C.surf,borderRadius:7,border:`1px solid ${C.bdr}`,padding:'7px 11px',marginBottom:12}}>
      <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="Search items in this category…" style={{width:'100%',border:'none',outline:'none',fontSize:13,color:C.txt,fontFamily:'inherit',background:'transparent'}}/>
    </div>
    <div style={{border:mobile?'none':`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden'}}>
      <CardTable mobile={mobile} columns={detailColumns} rows={pageRows} keyFn={(m,i)=>i} emptyLabel="No items match"/>
    </div>
    {totalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:12}}>
      <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {page+1} of {totalPages} · {itemRows.length} items</span>
      <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}
  </Modal>;
}

function CategoryAnalytics(){
  const[catDetail,setCatDetail]=useState(null);
  const rows=useMemo(()=>{
    const metrics=getAllItemMetrics();
    const expectedMap=getExpectedStockMap();
    // See getFifoStockMap in store.js — replaces the flat, easily-stale
    // item.purch×stock estimate with FIFO-costed layers built from real
    // dated purchase/sale/repair events.
    const fifoMap=getFifoStockMap();
    const grandRevenue=metrics.reduce((s,m)=>s+m.totalRevenue,0);
    const byCat={};
    for(const m of metrics){
      const cat=m.item.cat||'Other';
      if(!byCat[cat])byCat[cat]={cat,count:0,stock:0,stockValue:0,damaged:0,totalSoldQty:0,totalRevenue:0,profit:0,expected:0};
      const b=byCat[cat];
      const stock=m.item.stock||0;
      const fifo=fifoMap[m.item.id];
      b.count++; b.stock+=stock; b.stockValue+=fifo?fifo.value:(stock*(m.item.purch||0)); b.damaged+=(m.item.damaged||0);
      b.totalSoldQty+=m.totalSoldQty; b.totalRevenue+=m.totalRevenue; b.profit+=m.profit; b.expected+=(expectedMap[m.item.id]||0);
    }
    return Object.values(byCat).map(b=>({
      ...b,
      avgProfitPct: b.totalRevenue>0?Math.round(b.profit/b.totalRevenue*1000)/10:0,
      shareOfBusiness: grandRevenue>0?Math.round(b.totalRevenue/grandRevenue*1000)/10:0,
    })).sort((a,b)=>b.totalRevenue-a.totalRevenue);
  },[state.sales.length,state.purchases.length,state.items.length,state.repairExpenses.length]);
  const totals=useMemo(()=>rows.reduce((s,r)=>({count:s.count+r.count,stock:s.stock+r.stock,stockValue:s.stockValue+r.stockValue,damaged:s.damaged+r.damaged,totalRevenue:s.totalRevenue+r.totalRevenue,profit:s.profit+r.profit,expected:s.expected+r.expected}),{count:0,stock:0,stockValue:0,damaged:0,totalRevenue:0,profit:0,expected:0}),[rows]);
  const mobile=useIsMobile();
  const columns=[
    {key:'cat',label:'Category',render:r=><Badge v={r.cat} color="blue"/>},
    {key:'count',label:'Items',align:'right',render:r=><span style={{color:C.sub}}>{r.count}</span>},
    {key:'stock',label:'Stock (pcs)',align:'right',render:r=><span style={{color:C.sub}}>{r.stock}</span>},
    {key:'expected',label:'Expected',align:'right',render:r=><span style={{color:r.expected>0?C.amber:C.muted}}>{r.expected>0?'+'+r.expected:'—'}</span>},
    {key:'stockValue',label:'Stock Value',align:'right',render:r=><span style={{fontWeight:600,color:C.txt}}>{fmt(r.stockValue)}</span>},
    {key:'sold',label:'Sold (pcs)',align:'right',render:r=><span style={{color:C.sub}}>{r.totalSoldQty}</span>},
    {key:'revenue',label:'Revenue',align:'right',render:r=><span style={{fontWeight:600,color:C.txt}}>{fmt(r.totalRevenue)}</span>},
    {key:'profit',label:'Profit',align:'right',render:r=><span style={{fontWeight:600,color:r.profit>=0?C.blue:C.red}}>{fmt(r.profit)}</span>},
    {key:'profitPct',label:'Profit %',align:'right',render:r=><Badge v={r.avgProfitPct+'%'} color={r.avgProfitPct>25?'green':r.avgProfitPct>15?'amber':'red'}/>},
    {key:'share',label:'Share of Biz',align:'right',render:r=><span style={{color:C.sub}}>{r.shareOfBusiness}%</span>},
    {key:'damaged',label:'Damaged',align:'right',render:r=><span style={{color:r.damaged>0?C.amber:C.sub}}>{r.damaged}</span>},
  ];
  const footer=rows.length>0&&<div style={{display:'flex',flexWrap:'wrap',gap:14,justifyContent:'space-between',
    padding:mobile?'10px 14px':'9px 12px',background:'#F6F8FA',fontWeight:700,fontSize:13,
    borderTop:mobile?`1px solid ${C.bdr}`:'none',borderRadius:mobile?9:0}}>
    <span>Total</span>
    <span style={{display:'flex',gap:14,flexWrap:'wrap'}}>
      <span>{totals.count} items</span><span>{totals.stock} pcs</span><span>{fmt(totals.stockValue)}</span>
      <span>{fmt(totals.profit)} profit</span><span>{totals.damaged} damaged</span>
    </span>
  </div>;
  return <div>
    {catDetail&&<CategoryDetail cat={catDetail} onClose={()=>setCatDetail(null)}/>}
    <SectionHead title="Category Analytics" sub={`${totals.count} active items · ${fmt(totals.stockValue)} total stock value · click a category for item-level detail`}/>
    <div style={{background:mobile?'transparent':C.surf,borderRadius:8,border:mobile?'none':`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <CardTable mobile={mobile} columns={columns} rows={rows} keyFn={(r,i)=>i}
        onRowClick={r=>setCatDetail(r.cat)} footer={footer}/>
    </div>
  </div>;
}

function ItemLog(){
  const[search,setSearch]=useState('');
  const[dateFrom,setDateFrom]=useState('');
  const[dateTo,setDateTo]=useState('');
  const[dirF,setDirF]=useState('');
  const[page,setPage]=useState(0);
  const PAGE=50;

  const allRows=useMemo(()=>getAllItemRows().sort((a,b)=>b.d.localeCompare(a.d)),[state.sales.length,state.purchases.length]);

  const filtered=useMemo(()=>allRows.filter(r=>
    (!search||(r.item||'').toLowerCase().includes(search.toLowerCase())||(r.party||'').toLowerCase().includes(search.toLowerCase()))&&
    (!dateFrom||r.d>=dateFrom)&&(!dateTo||r.d<=dateTo)&&
    (!dirF||r.dir===dirF)
  ),[allRows,search,dateFrom,dateTo,dirF]);

  const pageData=filtered.slice(page*PAGE,(page+1)*PAGE);
  const totalPages=Math.ceil(filtered.length/PAGE);
  const totalIn=Math.round(filtered.filter(r=>r.dir==='in').reduce((s,r)=>s+(r.qty||0),0));
  const totalOut=Math.round(Math.abs(filtered.filter(r=>r.dir==='out').reduce((s,r)=>s+(r.qty||0),0)));
  const mobile=useIsMobile();
  const columns=[
    {key:'date',label:'Date',render:r=><span style={{whiteSpace:'nowrap',color:C.sub}}>{fmtD(r.d)}{r.isNew&&<span style={{fontSize:10,color:C.blue,fontWeight:700,marginLeft:5}}>NEW</span>}</span>},
    {key:'item',label:'Item',render:r=><span style={{fontWeight:600,color:C.txt}}>{r.item||'—'}</span>},
    {key:'party',label:'Party',render:r=><span style={{color:C.sub}}>{r.party||'—'}</span>},
    {key:'qty',label:'Qty',align:'right',render:r=><span style={{color:r.dir==='out'&&r.qty<0?C.amber:C.txt}}>{r.qty}</span>},
    {key:'rate',label:'Rate ₹',align:'right',render:r=><span style={{color:C.sub}}>{fmt(r.rate)}</span>},
    {key:'amount',label:'Amount',align:'right',render:r=><span style={{fontWeight:700,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(r.amount)}</span>},
    {key:'type',label:'Type',align:'center',render:r=><Badge v={r.isReturn?'Return':r.dir==='in'?'Purchased':'Sold'} color={r.isReturn?'amber':r.dir==='in'?'green':'blue'}/>},
  ];

  return <div>
    <SectionHead title="Item Log" sub={`${filtered.length} lines · ${totalIn} pcs in · ${totalOut} pcs out`}/>
    <div style={{display:'flex',gap:10,marginBottom:14,flexWrap:'wrap',alignItems:'center'}}>
      <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="Search item or party…"
        style={{flex:1,minWidth:200,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}
        onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
      <input type="date" value={dateFrom} onChange={e=>{setDateFrom(e.target.value);setPage(0);}}
        style={{padding:'8px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
      <input type="date" value={dateTo} onChange={e=>{setDateTo(e.target.value);setPage(0);}}
        style={{padding:'8px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
      <Chip v="All" active={!dirF} onClick={()=>{setDirF('');setPage(0);}}/>
      <Chip v="Purchased" active={dirF==='in'} onClick={()=>{setDirF('in');setPage(0);}}/>
      <Chip v="Sold" active={dirF==='out'} onClick={()=>{setDirF('out');setPage(0);}}/>
      {(search||dateFrom||dateTo||dirF)&&<Btn onClick={()=>{setSearch('');setDateFrom('');setDateTo('');setDirF('');setPage(0);}} variant="secondary" size="sm">Clear</Btn>}
    </div>
    <div style={{background:mobile?'transparent':C.surf,borderRadius:8,border:mobile?'none':`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <CardTable mobile={mobile} columns={columns} rows={pageData} keyFn={(r,i)=>i} emptyLabel="No entries match"
        rowStyle={r=>r.isNew?{background:C.blueSoft}:null}/>
    </div>
    {totalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:14,flexWrap:'wrap'}}>
      <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {page+1} of {totalPages} · {filtered.length} lines</span>
      <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}
  </div>;
}

function DamagedInventory(){
  const[items,setItems]=useState(()=>state.items.filter(it=>(it.damaged||0)>0));
  const[repairQty,setRepairQty]=useState({});
  const[fixQty,setFixQty]=useState({});
  const{toast,TC}=useToast();
  const refresh=()=>setItems(state.items.filter(it=>(it.damaged||0)>0));
  // A returned-goods trail (a sales return that touched this item) means the
  // damaged qty is legitimate. Anything sitting in damaged with NO such trail
  // can only have gotten there before purchases were fixed to land in stock
  // instead — i.e. it's very likely fresh stock stuck in the wrong bucket.
  const hasReturnHistory=useMemo(()=>{
    const ids=new Set();
    for(const s of state.sales) if(s.type==='return') for(const l of (s.items||[])) if(l.itemId) ids.add(l.itemId);
    return ids;
  },[state.sales.length]);
  // Who each item's damaged pieces were recovered from — cross-references
  // every return line landed as 'damaged' back to the return it came from, so
  // "why is this damaged" has an answer beyond just a running count.
  const damageSources=useMemo(()=>{
    const map={};
    for(const s of state.sales){
      if(s.type!=='return')continue;
      for(const l of (s.items||[])){
        if(!l.itemId||l.destination!=='damaged')continue;
        (map[l.itemId]=map[l.itemId]||[]).push({date:s.date,party:s.partyName,qty:Math.abs(l.qty||0),note:l.damageNote||''});
      }
    }
    for(const arr of Object.values(map))arr.sort((a,b)=>b.date.localeCompare(a.date));
    return map;
  },[state.sales.length]);
  const doRepair=it=>{
    const q=Number(repairQty[it.id])||0;
    if(q<=0||q>(it.damaged||0))return toast('Enter a valid qty (≤ damaged)','error');
    repairStock(it.id,q);
    setRepairQty(r=>({...r,[it.id]:''}));
    refresh();
    toast('Moved '+q+' pcs to repaired stock ✓');
  };
  const doCorrect=it=>{
    const q=Number(fixQty[it.id])||0;
    if(q<=0||q>(it.damaged||0))return toast('Enter a valid qty (≤ damaged)','error');
    correctDamagedToStock(it.id,q);
    setFixQty(f=>({...f,[it.id]:''}));
    refresh();
    toast('Corrected '+q+' pcs back to stock ✓');
  };
  const totalDamaged=items.reduce((s,it)=>s+(it.damaged||0),0);
  const totalValue=items.reduce((s,it)=>s+(it.damaged||0)*(it.purch||0),0);
  const mobile=useIsMobile();
  const columns=[
    {key:'item',label:'Item',render:it=>{
      const suspect=!hasReturnHistory.has(it.id);
      return <div>
        <span style={{fontWeight:600,color:C.txt}}>{it.internal||it.name}</span>
        {suspect&&<div title="No sales-return on record for this item — likely a purchase that landed here before that bug was fixed, not real damage" style={{fontSize:10,color:C.amberBright,fontWeight:700,marginTop:2}}>⚠ No return on record</div>}
      </div>;
    }},
    {key:'cat',label:'Category',render:it=><span style={{color:C.sub}}>{it.cat||'—'}</span>},
    {key:'damaged',label:'Damaged (pcs)',align:'right',render:it=><span style={{fontWeight:700,color:C.amber}}>{it.damaged||0}</span>},
    {key:'value',label:'Value',align:'right',render:it=><span style={{color:C.sub}}>{fmt((it.damaged||0)*(it.purch||0))}</span>},
    {key:'repair',label:'Repair (genuinely damaged)',align:'right',render:it=>
      <div style={{display:'flex',gap:4,justifyContent:'flex-end',alignItems:'center'}}>
        <input type="number" value={repairQty[it.id]||''} onChange={e=>setRepairQty(r=>({...r,[it.id]:e.target.value}))}
          placeholder="qty" onKeyDown={e=>{if(e.key==='Enter')doRepair(it);}}
          style={{width:55,padding:'4px 6px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',fontFamily:'inherit'}}/>
        <Btn onClick={()=>doRepair(it)} size="sm">Repair</Btn>
      </div>},
    {key:'correct',label:'Correct (never damaged)',align:'right',render:it=>
      <div style={{display:'flex',gap:4,justifyContent:'flex-end',alignItems:'center'}}>
        <input type="number" value={fixQty[it.id]||''} onChange={e=>setFixQty(f=>({...f,[it.id]:e.target.value}))}
          placeholder="qty" onKeyDown={e=>{if(e.key==='Enter')doCorrect(it);}}
          style={{width:55,padding:'4px 6px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',fontFamily:'inherit'}}/>
        <Btn onClick={()=>doCorrect(it)} size="sm" variant="secondary">→ Stock</Btn>
      </div>},
  ];
  const renderSubRow=it=>{
    const sources=damageSources[it.id]||[];
    if(!sources.length)return null;
    const suspect=!hasReturnHistory.has(it.id);
    return <div style={{padding:mobile?'8px 0 0':'2px 12px 9px 12px',background:mobile?'transparent':(suspect?C.amberSoft:'#FAFAFA')}}>
      <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.3,marginBottom:3}}>Recovered from</div>
      {sources.slice(0,4).map((s,i)=><div key={i} style={{display:'flex',justifyContent:'space-between',fontSize:11,color:C.sub,padding:'1px 0'}}>
        <span>{fmtD(s.date)} · {s.party} · {s.qty} pc{s.qty===1?'':'s'}{s.note?' — '+s.note:''}</span>
      </div>)}
      {sources.length>4&&<div style={{fontSize:10,color:C.muted}}>+{sources.length-4} more</div>}
    </div>;
  };
  return <div>
    <TC/>
    <SectionHead title="Damaged Inventory" sub={`${items.length} items · ${totalDamaged} pcs · ${fmt(totalValue)} value`}/>
    <div style={{background:mobile?'transparent':C.surf,borderRadius:8,border:mobile?'none':`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <CardTable mobile={mobile} columns={columns} rows={items} keyFn={it=>it.id} emptyLabel="No damaged stock ✓"
        renderSubRow={renderSubRow}
        rowStyle={it=>!hasReturnHistory.has(it.id)?{background:C.amberSoft}:null}/>
    </div>
  </div>;
}

function RepairedInventory(){
  const[items,setItems]=useState(()=>state.items.filter(it=>(it.repaired||0)>0));
  const[editVal,setEditVal]=useState({});
  const{toast,TC}=useToast();
  const refresh=()=>setItems(state.items.filter(it=>(it.repaired||0)>0));
  // Repaired stock is a direct override here rather than a transfer action —
  // unlike damaged/stock, there's no "move" counterpart to it yet (no sale
  // flow draws from repaired), so correcting a wrong count just means
  // setting the number, the same field Item Masters' edit form exposes.
  const doSet=it=>{
    const v=editVal[it.id];
    if(v===undefined||v==='')return;
    const n=Number(v);
    if(isNaN(n)||n<0)return toast('Enter a valid quantity','error');
    state.items=state.items.map(x=>x.id===it.id?{...x,repaired:n}:x);
    save.items();
    setEditVal(e=>({...e,[it.id]:''}));
    refresh();
    toast('Updated ✓');
  };
  const totalRepaired=items.reduce((s,it)=>s+(it.repaired||0),0);
  const totalValue=items.reduce((s,it)=>s+(it.repaired||0)*(it.purch||0),0);
  const mobile=useIsMobile();
  const columns=[
    {key:'item',label:'Item',render:it=><span style={{fontWeight:600,color:C.txt}}>{it.internal||it.name}</span>},
    {key:'cat',label:'Category',render:it=><span style={{color:C.sub}}>{it.cat||'—'}</span>},
    {key:'repaired',label:'Repaired (pcs)',align:'right',render:it=><span style={{fontWeight:700,color:C.green}}>{it.repaired||0}</span>},
    {key:'value',label:'Value',align:'right',render:it=><span style={{color:C.sub}}>{fmt((it.repaired||0)*(it.purch||0))}</span>},
    {key:'correct',label:'Correct qty',align:'right',render:it=>
      <div style={{display:'flex',gap:4,justifyContent:'flex-end',alignItems:'center'}}>
        <input type="number" value={editVal[it.id]??''} onChange={e=>setEditVal(v=>({...v,[it.id]:e.target.value}))}
          placeholder={String(it.repaired||0)} onKeyDown={e=>{if(e.key==='Enter')doSet(it);}}
          style={{width:60,padding:'4px 6px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',fontFamily:'inherit'}}/>
        <Btn onClick={()=>doSet(it)} size="sm" variant="secondary">Set</Btn>
      </div>},
  ];
  return <div>
    <TC/>
    <SectionHead title="Repaired Inventory" sub={`${items.length} items · ${totalRepaired} pcs · ${fmt(totalValue)} value`}/>
    <div style={{fontSize:12,color:C.sub,marginBottom:12}}>Kept separate from fresh stock — repaired units are for selling on their own (often at a discount), never mixed into a regular sale's available quantity.</div>
    <div style={{background:mobile?'transparent':C.surf,borderRadius:8,border:mobile?'none':`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <CardTable mobile={mobile} columns={columns} rows={items} keyFn={it=>it.id} emptyLabel="No repaired stock"/>
    </div>
  </div>;
}

function ItemNamingEditor(){
  const[items,setItems]=useState(()=>[...state.items].sort((a,b)=>(a.cat||'').localeCompare(b.cat||'')||(a.internal||a.name||'').localeCompare(b.internal||b.name||'')));
  const[search,setSearch]=useState('');
  const{toast,TC}=useToast();
  const filtered=items.filter(it=>!search||[it.internal,it.customer,it.supplier,it.dispatchName].some(v=>(v||'').toLowerCase().includes(search.toLowerCase())));
  const update=(id,field,val)=>setItems(list=>list.map(it=>it.id===id?{...it,[field]:val}:it));
  const commit=(id,field)=>{
    const it=items.find(x=>x.id===id);
    state.items=state.items.map(x=>x.id===id?{...x,[field]:it[field]}:x);
    save.items();
  };
  return <div>
    <TC/>
    <SectionHead title="Naming Conventions" sub={`${filtered.length} items · edit any name, click away to save`}/>
    <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search items…"
      style={{width:'100%',marginBottom:14,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}
      onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'auto'}}>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>
          {['Category','Internal Name','Customer Name','Supplier Name','Dispatch Name'].map((h,i)=><th key={i} style={{padding:'9px 12px',textAlign:'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`,whiteSpace:'nowrap'}}>{h}</th>)}
        </tr></thead>
        <tbody>
          {filtered.length===0&&<tr><td colSpan={5} style={{padding:'40px',textAlign:'center',color:C.muted}}>No items match</td></tr>}
          {filtered.map(it=>(
            <tr key={it.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
              <td style={{padding:'6px 12px',color:C.sub,whiteSpace:'nowrap'}}>{it.cat||'—'}{it.active===false?<span style={{marginLeft:6,fontSize:10,color:C.muted}}>(inactive)</span>:null}</td>
              {['internal','customer','supplier','dispatchName'].map(field=>(
                <td key={field} style={{padding:'4px 6px'}}>
                  <input value={it[field]||''} onChange={e=>update(it.id,field,e.target.value)} onBlur={()=>commit(it.id,field)}
                    style={{width:'100%',minWidth:140,padding:'6px 8px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12.5,outline:'none',fontFamily:'inherit'}}
                    onFocus={e=>e.target.style.borderColor=C.blue}/>
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  </div>;
}

function ItemDetail({item,onClose}){
  const{toast,TC}=useToast();
  const mobile=useIsMobile();
  const[search,setSearch]=useState('');
  const[page,setPage]=useState(0);
  const[damaged,setDamaged]=useState(item.damaged||0);
  const[repaired,setRepaired]=useState(item.repaired||0);
  const[repairQty,setRepairQty]=useState('');
  const[dmgQty,setDmgQty]=useState('');
  const[fixQty,setFixQty]=useState('');
  const[stock,setStock]=useState(item.stock||0);
  const PAGE=20;
  const metrics=useMemo(()=>getItemMetrics(item),[item.id,state.sales.length,state.purchases.length,state.items.length]);
  const grandRevenue=useMemo(()=>getAllItemMetrics().reduce((s,m)=>s+m.totalRevenue,0),[state.sales.length,state.purchases.length]);
  const shareOfBusiness=grandRevenue>0?Math.round(metrics.totalRevenue/grandRevenue*1000)/10:0;
  const peti=item.peti||1;
  const petiCount=Math.floor(stock/peti);
  const looseCount=stock-petiCount*peti;
  // FIFO-costed value/avg rate (see getFifoStockMap in store.js) instead of
  // the flat item.purch×stock estimate — item.purch only updates when the
  // offline seed script re-runs, so it drifts stale as soon as an in-app
  // purchase lands at a different rate.
  const fifo=useMemo(()=>getFifoStockMap()[item.id],[item.id,state.purchases.length,state.sales.length,state.repairExpenses.length,state.items.length]);
  const stockValue=fifo?fifo.value:(stock*(item.purch||0));
  const fifoAvgRate=fifo?Math.round(fifo.avgRate*100)/100:(item.purch||0);
  const expectedQty=useMemo(()=>state.purchases.filter(p=>!p.receivedDate).reduce((s,p)=>{
    const l=(p.items||[]).find(x=>x.itemId===item.id);
    return s+(l?l.qty:0);
  },0),[item.id,state.purchases.length]);
  const rows=useMemo(()=>metrics.rows.filter(r=>!search||(r.party||'').toLowerCase().includes(search.toLowerCase())).sort((a,b)=>b.d.localeCompare(a.d)),[metrics,search]);
  const pageRows=rows.slice(page*PAGE,(page+1)*PAGE);
  const totalPages=Math.ceil(rows.length/PAGE);
  const doRepair=()=>{
    const q=Number(repairQty)||0;
    if(q<=0||q>damaged)return toast('Enter a valid qty (≤ damaged)','error');
    repairStock(item.id,q);
    setDamaged(d=>d-q); setRepaired(r=>r+q); setRepairQty('');
    toast('Moved '+q+' pcs to repaired stock ✓');
  };
  const doMoveDamaged=()=>{
    const q=Number(dmgQty)||0;
    if(q<=0||q>stock)return toast('Enter a valid qty (≤ stock)','error');
    moveToDamaged(item.id,q);
    setStock(s=>s-q); setDamaged(d=>d+q); setDmgQty('');
    toast('Moved '+q+' pcs to damaged');
  };
  // Not a repair — for units that were never actually damaged (e.g. a
  // purchase mis-routed into the damaged bucket by a since-fixed bug), this
  // corrects them straight back to plain stock without touching repaired.
  const doCorrect=()=>{
    const q=Number(fixQty)||0;
    if(q<=0||q>damaged)return toast('Enter a valid qty (≤ damaged)','error');
    correctDamagedToStock(item.id,q);
    setDamaged(d=>d-q); setStock(s=>s+q); setFixQty('');
    toast('Corrected '+q+' pcs back to stock ✓');
  };
  const Card=({label,value,color})=>(
    <div style={{background:'#F9FAFB',borderRadius:7,padding:'8px 10px',border:`1px solid ${C.bdr}`}}>
      <div style={{fontSize:9,color:C.sub,textTransform:'uppercase',letterSpacing:.4}}>{label}</div>
      <div style={{fontSize:14,fontWeight:800,color:color||C.txt}}>{value}</div>
    </div>
  );
  return <Modal title={(item.internal||item.name)+' — Item Detail'} onClose={onClose} width={680}>
    <TC/>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:16}}>
      <div>
        <div style={{fontSize:11,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>Current Stock</div>
        <div style={{fontSize:26,fontWeight:800,color:stock<=0?C.red:C.txt,fontVariantNumeric:'tabular-nums'}}>{petiCount} peti + {looseCount} pcs</div>
        <div style={{fontSize:11,color:C.sub,marginTop:2}}>{stock} pcs total · {fmt(stockValue)} value</div>
      </div>
      <div style={{textAlign:'right'}}>
        <Badge v={item.cat||'Other'} color="blue"/>
        <div style={{fontSize:11,color:C.sub,marginTop:6}}>Sale ₹{item.sale} · Purch ₹{item.purch}{fifo&&Math.abs(fifoAvgRate-(item.purch||0))>0.01?<span style={{color:C.blue}}> · FIFO avg ₹{fifoAvgRate}</span>:null}</div>
      </div>
    </div>
    <div style={{display:'grid',gridTemplateColumns:mobile?'repeat(2,1fr)':'repeat(4,1fr)',gap:8,marginBottom:16}}>
      <Card label="Total Sold" value={metrics.totalSoldQty+' pcs'}/>
      <Card label="Total Purchased" value={metrics.totalPurchQty+' pcs'}/>
      <Card label="Profit Earned" value={fmt(metrics.profit)} color={metrics.profit>=0?C.blue:C.red}/>
      <Card label="Avg Profit %" value={metrics.avgProfitPct+'%'}/>
      <Card label="Total Revenue" value={fmt(metrics.totalRevenue)}/>
      <Card label="Share of Business" value={shareOfBusiness+'%'}/>
      <div style={{background:'#F9FAFB',borderRadius:7,padding:'8px 10px',border:`1px solid ${C.bdr}`}}>
        <div style={{fontSize:9,color:C.sub,textTransform:'uppercase',letterSpacing:.4}}>Damaged Stock</div>
        <div style={{fontSize:14,fontWeight:800,color:damaged>0?C.amber:C.txt}}>{damaged} pcs</div>
        {stock>0&&<div style={{display:'flex',gap:4,alignItems:'center',marginTop:4}}>
          <input type="number" value={dmgQty} onChange={e=>setDmgQty(e.target.value)} placeholder="qty"
            onKeyDown={e=>{if(e.key==='Enter')doMoveDamaged();}}
            style={{width:50,padding:'2px 5px',borderRadius:4,border:`1px solid ${C.bdr}`,fontSize:12,fontFamily:'inherit'}}/>
          <button onClick={doMoveDamaged} style={{border:'none',background:'none',color:C.amber,cursor:'pointer',fontSize:11,fontWeight:700}}>← Move here</button>
        </div>}
        {damaged>0&&<div style={{display:'flex',gap:4,alignItems:'center',marginTop:4}}>
          <input type="number" value={fixQty} onChange={e=>setFixQty(e.target.value)} placeholder="qty"
            onKeyDown={e=>{if(e.key==='Enter')doCorrect();}}
            style={{width:50,padding:'2px 5px',borderRadius:4,border:`1px solid ${C.bdr}`,fontSize:12,fontFamily:'inherit'}}/>
          <button onClick={doCorrect} title="For units that were never actually damaged — e.g. a purchase wrongly landed here"
            style={{border:'none',background:'none',color:C.muted,cursor:'pointer',fontSize:11,fontWeight:700,textDecoration:'underline'}}>Not damaged, → stock</button>
        </div>}
      </div>
      <div style={{background:'#F9FAFB',borderRadius:7,padding:'8px 10px',border:`1px solid ${C.bdr}`}}>
        <div style={{fontSize:9,color:C.sub,textTransform:'uppercase',letterSpacing:.4}}>Repaired Stock</div>
        <div style={{fontSize:14,fontWeight:800,color:repaired>0?C.green:C.txt}}>{repaired} pcs</div>
        {damaged>0&&<div style={{display:'flex',gap:4,alignItems:'center',marginTop:4}}>
          <input type="number" value={repairQty} onChange={e=>setRepairQty(e.target.value)} placeholder="qty"
            onKeyDown={e=>{if(e.key==='Enter')doRepair();}}
            style={{width:50,padding:'2px 5px',borderRadius:4,border:`1px solid ${C.bdr}`,fontSize:12,fontFamily:'inherit'}}/>
          <button onClick={doRepair} style={{border:'none',background:'none',color:C.blue,cursor:'pointer',fontSize:11,fontWeight:700}}>Repair</button>
        </div>}
      </div>
      <Card label="Stock Value" value={fmt(stockValue)}/>
      <Card label="Expected (Incoming)" value={expectedQty+' pcs'} color={expectedQty>0?C.amber:C.txt}/>
    </div>
    <div style={{background:C.surf,borderRadius:7,border:`1px solid ${C.bdr}`,padding:'7px 11px',marginBottom:12}}>
      <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="Search by party…" style={{width:'100%',border:'none',outline:'none',fontSize:13,color:C.txt,fontFamily:'inherit',background:'transparent'}}/>
    </div>
    <div style={{border:mobile?'none':`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden'}}>
      <CardTable mobile={mobile} keyFn={(r,i)=>i} rows={pageRows} emptyLabel="No purchase/sale entries for this item"
        columns={[
          {key:'date',label:'Date',render:r=><span style={{whiteSpace:'nowrap',color:C.sub}}>{fmtD(r.d)}</span>},
          {key:'party',label:'Party',render:r=><span style={{fontWeight:500,color:C.txt}}>{r.party||'—'}</span>},
          {key:'qty',label:'Qty',align:'right',render:r=>Math.abs(r.qty||0)},
          {key:'rate',label:'Rate',align:'right',render:r=><span style={{color:C.sub}}>{fmt(r.rate)}</span>},
          {key:'amount',label:'Amount',align:'right',render:r=><span style={{fontWeight:700,color:C.blue}}>{fmt(Math.abs(r.amount||0))}</span>},
          {key:'type',label:'Type',align:'center',render:r=><Badge v={r.isReturn?'Return':r.dir==='in'?'Purchased':'Sold'} color={r.isReturn?'amber':r.dir==='in'?'green':'blue'}/>},
        ]}/>
    </div>
    {totalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:12}}>
      <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {page+1} of {totalPages} · {rows.length} entries</span>
      <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}
  </Modal>;
}

function ItemMasters(){
  const[items,setItems]=useState(state.items);
  const[search,setSearch]=useState('');
  const[catF,setCatF]=useState('');
  const[statusF,setStatusF]=useState('active');
  const[page,setPage]=useState(0);
  const[modal,setModal]=useState(null);
  const[form,setForm]=useState({});
  const[view,setView]=useState('list');
  const[detailItem,setDetailItem]=useState(null);
  const[mergeSrc,setMergeSrc]=useState(null);
  const[moveSrc,setMoveSrc]=useState(null);
  const[editSaleId,setEditSaleId]=useState(null);
  const[editSaleVal,setEditSaleVal]=useState('');
  const{toast,TC}=useToast();
  const mobile=useIsMobile();
  const PAGE=50;
  const cats=useMemo(()=>[...new Set(state.items.map(i=>i.cat))].sort(),[]);
  const catOptions=useMemo(()=>[...new Set([...cats,'Other'])].sort(),[cats]);
  const expectedMap=useMemo(()=>getExpectedStockMap(),[state.purchases.length]);
  // Purch ₹ column: FIFO avg cost of the item's current balance (see
  // getFifoStockMap in store.js) instead of the flat item.purch field, which
  // only ever changes when someone hand-edits it — it drifts stale the
  // moment an in-app purchase lands at a different rate.
  const fifoMap=useMemo(()=>getFifoStockMap(),[state.purchases.length,state.sales.length,state.repairExpenses.length,state.items.length]);
  const purchOf=it=>{ const f=fifoMap[it.id]; return f&&f.qty>0?Math.round(f.avgRate*100)/100:(it.purch||0); };
  // Sale ₹ column: item.sale when someone has actually set it, else the mode
  // of the item's recent in-app sale rates ("what we're actually charging
  // lately") — auto-suggested until manually overridden inline in this table.
  const saleOf=it=>(it.sale>0?it.sale:(getRecentModeSaleRate(it.id)??0));
  const commitSale=(it,raw)=>{
    const v=Math.max(0,Number(raw)||0);
    state.items=state.items.map(x=>x.id===it.id?{...x,sale:v}:x);
    save.items();
    setItems([...state.items]);
    setEditSaleId(null);
  };
  const filtered=items.filter(it=>
    (statusF==='active'?it.active!==false:it.active===false)&&
    (!catF||it.cat===catF)&&
    (!search||(it.customer||it.internal||'').toLowerCase().includes(search.toLowerCase()))
  ).sort((a,b)=>(a.cat||'').localeCompare(b.cat||'')||(a.sortOrder??Infinity)-(b.sortOrder??Infinity)||(a.internal||'').localeCompare(b.internal||''));
  const pageData=filtered.slice(page*PAGE,(page+1)*PAGE);
  const totalPages=Math.ceil(filtered.length/PAGE);
  const openEdit=it=>{ setForm({...it}); setModal('edit'); };
  const doSave=()=>{
    if(!form.internal?.trim())return toast('Name required','error');
    if(modal==='add')state.items=[...state.items,{...form,stock:Number(form.stock)||0,damaged:Number(form.damaged)||0,repaired:Number(form.repaired)||0,active:true,id:uid('ITM')}];
    else state.items=state.items.map(x=>x.id===form.id?{...form,stock:Number(form.stock)||0,damaged:Number(form.damaged)||0,repaired:Number(form.repaired)||0}:x);
    save.items();
    setItems([...state.items]); toast(modal==='add'?'Item added':'Updated'); setModal(null);
  };
  const csvCell=v=>{
    const s=String(v??'');
    return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;
  };
  const downloadReport=()=>{
    const header=['Internal Name','Customer Name','Category','Peti Size','Sale Rate','Purchase Rate','Margin %','Stock (pcs)','Damaged (pcs)','Repaired (pcs)','Expected Incoming (pcs)','Stock Value (₹)','Status'];
    const rows=filtered.map(it=>{
      const sale=saleOf(it), purch=purchOf(it);
      const margin=sale>0?Math.round((sale-purch)/sale*100):0;
      const stock=it.stock||0;
      const value=fifoMap[it.id]&&fifoMap[it.id].qty>0?fifoMap[it.id].value:stock*purch;
      return [it.internal||it.name||'',it.customer||'',it.cat||'Other',it.peti||0,sale,purch,margin,stock,it.damaged||0,it.repaired||0,expectedMap[it.id]||0,value,it.active!==false?'Active':'Discontinued'];
    });
    const csv=[header,...rows].map(r=>r.map(csvCell).join(',')).join('\n');
    const blob=new Blob([csv],{type:'text/csv'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url; a.download='inventory-report-'+today()+'.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    toast(rows.length+' items exported ✓');
  };
  const downloadReportPDF=()=>{
    const rows=filtered.map(it=>{
      const sale=saleOf(it), purch=purchOf(it);
      const margin=sale>0?Math.round((sale-purch)/sale*100):0;
      const stock=it.stock||0;
      const value=fifoMap[it.id]&&fifoMap[it.id].qty>0?fifoMap[it.id].value:stock*purch;
      return {name:it.internal||it.name||'',cat:it.cat||'Other',peti:it.peti||0,sale,purch,margin,stock,damaged:it.damaged||0,repaired:it.repaired||0,expected:expectedMap[it.id]||0,value,status:it.active!==false?'Active':'Discontinued'};
    });
    const totalValue=rows.reduce((s,r)=>s+r.value,0);
    const rs=n=>'Rs '+Number(n||0).toLocaleString('en-IN');
    const filterBits=[statusF==='active'?'Active items':'Discontinued items',catF||null,search?'Search: "'+search+'"':null].filter(Boolean).join(' - ');
    const doc=new jsPDF({orientation:'landscape',unit:'pt',format:'a4'});
    doc.setFontSize(16); doc.setFont(undefined,'bold');
    doc.text('Upkaran ERP - Inventory Report',40,40);
    doc.setFontSize(10); doc.setFont(undefined,'normal');
    doc.text('Generated '+fmtD(today())+'   '+filterBits+'   '+rows.length+' items',40,58);
    autoTable(doc,{
      startY:72,
      head:[['Item','Category','Peti','Sale Rate','Purch Rate','Margin','Stock','Damaged','Repaired','Expected','Value','Status']],
      body:rows.map(r=>[r.name,r.cat,r.peti,rs(r.sale),rs(r.purch),r.margin+'%',r.stock,r.damaged||'-',r.repaired||'-',r.expected||'-',rs(r.value),r.status]),
      foot:[['','','','','','','','','','Total',rs(totalValue),'']],
      styles:{fontSize:8,cellPadding:4},
      headStyles:{fillColor:[37,99,235],textColor:255,fontStyle:'bold'},
      footStyles:{fillColor:[243,244,246],textColor:20,fontStyle:'bold'},
      columnStyles:{2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'},6:{halign:'right'},7:{halign:'right'},8:{halign:'right'},9:{halign:'right'},10:{halign:'right'}},
      margin:{left:40,right:40},
    });
    doc.save('inventory-report-'+today()+'.pdf');
    toast(rows.length+' items exported ✓');
  };
  return <div>
    <TC/>
    {detailItem&&<ItemDetail item={detailItem} onClose={()=>setDetailItem(null)}/>}
    {mergeSrc&&<MergeItemModal source={mergeSrc} onClose={()=>setMergeSrc(null)} onMerged={()=>{setItems([...state.items]);setMergeSrc(null);}}/>}
    {moveSrc&&<MoveStockModal source={moveSrc} onClose={()=>setMoveSrc(null)} onMoved={()=>{setItems([...state.items]);setMoveSrc(null);}}/>}
    <SectionHead title="Inventory" sub={`${items.filter(it=>it.active!==false).length} active items`} action={view==='list'?<div style={{display:'flex',gap:8}}><Btn onClick={downloadReport} variant="secondary" icon="⬇">CSV</Btn><Btn onClick={downloadReportPDF} variant="secondary" icon="⬇">PDF</Btn><Btn onClick={()=>{setForm({cat:'Other',peti:1,sale:0,purch:0,active:true});setModal('add');}}>+ Add Item</Btn></div>:null}/>
    <div style={{display:'flex',gap:8,marginBottom:16,flexWrap:'wrap'}}>
      <Chip v="Item List" active={view==='list'} onClick={()=>setView('list')}/>
      <Chip v="Category Analytics" active={view==='analytics'} onClick={()=>setView('analytics')}/>
      <Chip v="Item Log" active={view==='log'} onClick={()=>setView('log')}/>
      <Chip v="Naming Conventions" active={view==='naming'} onClick={()=>setView('naming')}/>
      <Chip v="Damaged Inventory" active={view==='damaged'} onClick={()=>setView('damaged')}/>
      <Chip v="Repaired Inventory" active={view==='repaired'} onClick={()=>setView('repaired')}/>
    </div>
    {view==='analytics'?<CategoryAnalytics/>:view==='log'?<ItemLog/>:view==='naming'?<ItemNamingEditor/>:view==='damaged'?<DamagedInventory/>:view==='repaired'?<RepairedInventory/>:<>
    {modal&&<Modal title={modal==='add'?'Add Item':'Edit Item'} onClose={()=>setModal(null)}>
      <Grid2>
        <div style={{gridColumn:'1/-1'}}><Inp label="Internal Name (Sales Sheet)" value={form.internal||''} onChange={v=>setForm(f=>({...f,internal:v,name:v}))} required hint="How it appears in your Excel/bills"/></div>
        <Inp label="Customer Name" value={form.customer||''} onChange={v=>setForm(f=>({...f,customer:v}))} hint="What you tell the buyer"/>
        <Inp label="Supplier Name" value={form.supplier||''} onChange={v=>setForm(f=>({...f,supplier:v}))} hint="Vendor catalog name"/>
        <Inp label="Dispatch Name" value={form.dispatchName||''} onChange={v=>setForm(f=>({...f,dispatchName:v}))} hint="What the warehouse/dispatch team calls it"/>
        <div>
          <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>Category</label>
          <select value={catOptions.includes(form.cat)?form.cat:'__new__'}
            onChange={e=>setForm(f=>({...f,cat:e.target.value==='__new__'?'':e.target.value}))}
            style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',background:C.surf,boxSizing:'border-box'}}>
            {catOptions.map(c=><option key={c} value={c}>{c}</option>)}
            <option value="__new__">+ New category…</option>
          </select>
          {!catOptions.includes(form.cat)&&<input value={form.cat||''} onChange={e=>setForm(f=>({...f,cat:e.target.value}))}
            placeholder="New category name" autoFocus
            style={{width:'100%',marginTop:6,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}/>}
        </div>
        <Inp label="Peti Size (pcs)" value={form.peti||1} onChange={v=>setForm(f=>({...f,peti:Number(v)}))} type="number"/>
        <Inp label="Sale Rate ₹/pc" value={form.sale||0} onChange={v=>setForm(f=>({...f,sale:Number(v)}))} type="number"/>
        <Inp label="Purchase Rate ₹/pc" value={form.purch||0} onChange={v=>setForm(f=>({...f,purch:Number(v)}))} type="number"/>
        <Inp label="Stock on Hand (pcs)" value={form.stock||0} onChange={v=>setForm(f=>({...f,stock:Number(v)}))} type="number" hint="Current qty available"/>
        <Inp label="Damaged Stock (pcs)" value={form.damaged||0} onChange={v=>setForm(f=>({...f,damaged:Number(v)}))} type="number" hint="Unsellable / spoiled qty"/>
        <Inp label="Repaired Stock (pcs)" value={form.repaired||0} onChange={v=>setForm(f=>({...f,repaired:Number(v)}))} type="number" hint="Fixed from damaged, tracked separately"/>
      </Grid2>
      <label style={{display:'flex',alignItems:'center',gap:8,marginTop:16,fontSize:13,color:C.txt,cursor:'pointer',fontWeight:600}}>
        <input type="checkbox" checked={form.active!==false} onChange={e=>setForm(f=>({...f,active:e.target.checked}))}/>
        Active {form.active===false&&<span style={{color:C.muted,fontWeight:400}}>— hidden from sale/purchase search and inventory reorder alerts</span>}
      </label>
      <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:20}}>
        <Btn onClick={()=>setModal(null)} variant="secondary">Cancel</Btn>
        <Btn onClick={doSave}>{modal==='add'?'Add Item':'Save Changes'}</Btn>
      </div>
    </Modal>}
    <div style={{display:'flex',gap:8,marginBottom:10,flexWrap:'wrap'}}>
      <Chip v="Active" active={statusF==='active'} onClick={()=>{setStatusF('active');setPage(0);}}/>
      <Chip v="Discontinued" active={statusF==='discontinued'} onClick={()=>{setStatusF('discontinued');setPage(0);}}/>
    </div>
    <div style={{display:'flex',gap:8,marginBottom:14,flexWrap:'wrap'}}>
      <Chip v="All" active={!catF} onClick={()=>{setCatF('');setPage(0);}}/>
      {cats.map(c=><Chip key={c} v={c} active={catF===c} onClick={()=>{setCatF(catF===c?'':c);setPage(0);}}/>)}
    </div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'8px 12px',marginBottom:12}}>
      <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="Search items…" style={{width:'100%',border:'none',outline:'none',fontSize:13,color:C.txt,fontFamily:'inherit',background:'transparent'}}/>
    </div>
    <div style={{background:mobile?'transparent':C.surf,borderRadius:8,border:mobile?'none':`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <CardTable mobile={mobile} rows={pageData} keyFn={(it,i)=>it.id||i} emptyLabel="No items match"
        groupHeader={(it,prev)=>{
          const prevCat=prev?prev.cat:(page>0?'__prevpage__':null);
          const showCatHeader=!catF&&(it.cat||'Other')!==(prevCat||'Other');
          const prevSubHead=prev?prev.subHead:(page>0?'__prevpage__':null);
          const showSubHead=it.subHead&&it.subHead!==prevSubHead;
          if(!showCatHeader&&!showSubHead)return null;
          return <>
            {showCatHeader&&<div style={{padding:'10px 12px',background:C.blueSoft,fontSize:13,fontWeight:800,color:C.blue,
              borderTop:prev?`2px solid ${C.blue}`:'none',borderRadius:mobile?7:0,marginTop:mobile&&prev?10:0}}>{it.cat||'Other'}</div>}
            {showSubHead&&<div style={{padding:'6px 12px',background:'#EEF2F6',fontSize:11,fontWeight:700,color:C.sub,
              textTransform:'uppercase',letterSpacing:.5,borderRadius:mobile?7:0,marginTop:mobile?6:0}}>{it.subHead}</div>}
          </>;
        }}
        columns={[
          {key:'internal',label:'Internal Name',render:it=><span onClick={()=>setDetailItem(it)}
            style={{fontWeight:600,color:C.blue,cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted'}}>{it.internal||it.name}</span>},
          {key:'customer',label:'Customer Name',cardLabel:'Customer',render:it=><span style={{fontSize:12,color:C.sub}}>{it.customer||'—'}</span>},
          {key:'cat',label:'Category',render:it=><Badge v={it.cat||'Other'} color="blue"/>},
          {key:'peti',label:'Peti',align:'right',render:it=><span style={{color:C.sub}}>{it.peti}</span>},
          {key:'sale',label:'Sale ₹',align:'right',render:it=>{
            const saleDisp=saleOf(it);
            return editSaleId===it.id?
              <input autoFocus type="number" value={editSaleVal}
                onChange={e=>setEditSaleVal(e.target.value)}
                onBlur={()=>commitSale(it,editSaleVal)}
                onKeyDown={e=>{if(e.key==='Enter')commitSale(it,editSaleVal); if(e.key==='Escape')setEditSaleId(null);}}
                style={{width:80,padding:'3px 6px',borderRadius:5,border:`1.5px solid ${C.blue}`,fontSize:13,textAlign:'right',outline:'none',fontFamily:'inherit'}}/>
              :<span onClick={()=>{setEditSaleId(it.id);setEditSaleVal(String(saleDisp||''));}}
                title={it.sale>0?'Click to edit':'Auto-suggested from recent sales · click to set manually'}
                style={{fontWeight:600,color:C.blue,cursor:'pointer',fontStyle:it.sale>0?'normal':'italic',borderBottom:'1px dotted currentColor'}}>{fmt(saleDisp)}</span>;
          }},
          {key:'purch',label:'Purch ₹',align:'right',render:it=><span style={{color:C.sub}} title="FIFO avg cost of current balance">{fmt(purchOf(it))}</span>},
          {key:'margin',label:'Margin',align:'right',render:it=>{
            const saleDisp=saleOf(it),purchDisp=purchOf(it);
            const margin=saleDisp>0?Math.round((saleDisp-purchDisp)/saleDisp*100):0;
            return <Badge v={margin+'%'} color={margin>25?'green':margin>15?'amber':'red'}/>;
          }},
          {key:'stock',label:'Stock',align:'right',render:it=>{
            const stock=it.stock||0;
            return <Badge v={stock+' pc'+(stock===1?'':'s')} color={stock<=0?'red':'gray'}/>;
          }},
          {key:'damaged',label:'Damaged',align:'right',render:it=>(it.damaged||0)>0?<Badge v={it.damaged+' pc'+(it.damaged===1?'':'s')} color="amber"/>:<span style={{color:C.muted}}>—</span>},
          {key:'expected',label:'Expected',align:'right',render:it=>{
            const expected=expectedMap[it.id]||0;
            return <span style={{color:expected>0?C.amber:C.muted,fontWeight:expected>0?700:400}}>{expected>0?'+'+expected:'—'}</span>;
          }},
          {key:'actions',label:'',cardLabel:'Actions',render:it=>
            <div style={{display:'flex',gap:6,justifyContent:'flex-end',flexWrap:'wrap'}}>
              <Btn onClick={()=>setMoveSrc(it)} variant="ghost" size="sm">Move</Btn>
              <Btn onClick={()=>setMergeSrc(it)} variant="ghost" size="sm">Merge</Btn>
              <Btn onClick={()=>openEdit(it)} variant="ghost" size="sm">Edit</Btn>
            </div>},
        ]}/>
    </div>
    {totalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:14}}>
      <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {page+1} of {totalPages} · {filtered.length} items</span>
      <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}
    </>}
  </div>;
}

function MergePartyModal({source,onClose,onMerged}){
  const[pq,setPq]=useState('');
  const[targetId,setTargetId]=useState('');
  const mobile=useIsMobile();
  const{toast,TC}=useToast();
  const candidates=useMemo(()=>state.parties.filter(p=>p.type===source.type&&p.id!==source.id&&(!pq||p.name.toLowerCase().includes(pq.toLowerCase()))).slice(0,8),[pq,source.id,source.type]);
  const target=state.parties.find(p=>p.id===targetId);
  const doMerge=()=>{
    if(!targetId)return toast('Pick a party to merge into','error');
    mergeParties(source.id,targetId);
    toast('Merged "'+source.name+'" into "'+target.name+'" ✓');
    onMerged?.();
  };
  return <Modal title={'Merge "'+source.name+'" into…'} onClose={onClose} width={440}>
    <TC/>
    <div style={{fontSize:12,color:C.sub,marginBottom:12}}>All sales, payments and follow-ups move to the target party, opening balances add up, and "{source.name}" is then removed. This can't be undone.</div>
    {!target
      ?<>
        <input value={pq} onChange={e=>setPq(e.target.value)} placeholder="Search party to merge into…" autoFocus
          style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit',marginBottom:8,boxSizing:'border-box'}}/>
        {pq.length>0&&candidates.map(p=><div key={p.id} onClick={()=>setTargetId(p.id)} style={{padding:'8px 12px',borderRadius:7,cursor:'pointer',border:`1px solid ${C.bdr}`,marginBottom:6}}
          {...(mobile?{}:{onMouseEnter:e=>e.currentTarget.style.background='#F6F8FA',onMouseLeave:e=>e.currentTarget.style.background='transparent'})}>
          <div style={{fontWeight:600,color:C.txt,fontSize:13}}>{p.name}</div>
          <div style={{fontSize:11,color:C.sub}}>{p.city||'—'}</div>
        </div>)}
      </>
      :<div style={{padding:'10px 14px',borderRadius:8,border:`1.5px solid ${C.blue}`,background:C.blueSoft,marginBottom:14}}>
        <div style={{fontSize:11,color:C.sub}}>Merging into</div>
        <div style={{fontWeight:700,color:C.txt,fontSize:14}}>{target.name}</div>
        <button onClick={()=>setTargetId('')} style={{marginTop:4,background:'none',border:'none',color:C.blue,cursor:'pointer',fontSize:11,textDecoration:'underline',fontFamily:'inherit'}}>Change</button>
      </div>}
    <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:10}}>
      <Btn onClick={onClose} variant="secondary">Cancel</Btn>
      <Btn onClick={doMerge} disabled={!targetId} variant="danger">Merge</Btn>
    </div>
  </Modal>;
}

function MergeItemModal({source,onClose,onMerged}){
  const[pq,setPq]=useState('');
  const[targetId,setTargetId]=useState('');
  const mobile=useIsMobile();
  const{toast,TC}=useToast();
  const candidates=useMemo(()=>state.items.filter(it=>it.id!==source.id&&it.active!==false&&(!pq||(it.internal||it.name||'').toLowerCase().includes(pq.toLowerCase()))).slice(0,8),[pq,source.id]);
  const target=state.items.find(it=>it.id===targetId);
  const doMerge=()=>{
    if(!targetId)return toast('Pick an item to merge into','error');
    mergeItems(source.id,targetId);
    toast('Merged "'+(source.internal||source.name)+'" into "'+(target.internal||target.name)+'" ✓');
    onMerged?.();
  };
  return <Modal title={'Merge "'+(source.internal||source.name)+'" into…'} onClose={onClose} width={440}>
    <TC/>
    <div style={{fontSize:12,color:C.sub,marginBottom:12}}>Stock, damaged and repaired quantities add to the target item, every sale/purchase line moves with them, and "{source.internal||source.name}" is then removed. This can't be undone.</div>
    {!target
      ?<>
        <input value={pq} onChange={e=>setPq(e.target.value)} placeholder="Search item to merge into…" autoFocus
          style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit',marginBottom:8,boxSizing:'border-box'}}/>
        {pq.length>0&&candidates.map(it=><div key={it.id} onClick={()=>setTargetId(it.id)} style={{padding:'8px 12px',borderRadius:7,cursor:'pointer',border:`1px solid ${C.bdr}`,marginBottom:6}}
          {...(mobile?{}:{onMouseEnter:e=>e.currentTarget.style.background='#F6F8FA',onMouseLeave:e=>e.currentTarget.style.background='transparent'})}>
          <div style={{fontWeight:600,color:C.txt,fontSize:13}}>{it.internal||it.name}</div>
          <div style={{fontSize:11,color:C.sub}}>{it.cat||'—'} · {it.stock||0} pcs in stock</div>
        </div>)}
      </>
      :<div style={{padding:'10px 14px',borderRadius:8,border:`1.5px solid ${C.blue}`,background:C.blueSoft,marginBottom:14}}>
        <div style={{fontSize:11,color:C.sub}}>Merging into</div>
        <div style={{fontWeight:700,color:C.txt,fontSize:14}}>{target.internal||target.name}</div>
        <div style={{fontSize:11,color:C.sub,marginTop:2}}>{target.stock||0} pcs → {(target.stock||0)+(source.stock||0)} pcs after merge</div>
        <button onClick={()=>setTargetId('')} style={{marginTop:4,background:'none',border:'none',color:C.blue,cursor:'pointer',fontSize:11,textDecoration:'underline',fontFamily:'inherit'}}>Change</button>
      </div>}
    <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:10}}>
      <Btn onClick={onClose} variant="secondary">Cancel</Btn>
      <Btn onClick={doMerge} disabled={!targetId} variant="danger">Merge</Btn>
    </div>
  </Modal>;
}

// Non-destructive counterpart to MergeItemModal: relocates a chosen quantity
// of stock to another item and leaves both items — and all their history —
// otherwise untouched. Used when two entries are genuinely distinct SKUs and
// only a stock count needs to move, not their identities.
function MoveStockModal({source,onClose,onMoved}){
  const[pq,setPq]=useState('');
  const[targetId,setTargetId]=useState('');
  const[qty,setQty]=useState(String(source.stock||0));
  const mobile=useIsMobile();
  const{toast,TC}=useToast();
  const candidates=useMemo(()=>state.items.filter(it=>it.id!==source.id&&it.active!==false&&(!pq||(it.internal||it.name||'').toLowerCase().includes(pq.toLowerCase()))).slice(0,8),[pq,source.id]);
  const target=state.items.find(it=>it.id===targetId);
  const q=Number(qty)||0;
  const doMove=()=>{
    if(!targetId)return toast('Pick an item to move stock into','error');
    if(q<=0||q>(source.stock||0))return toast('Enter a valid qty (≤ '+(source.stock||0)+')','error');
    moveItemStock(source.id,targetId,q);
    toast('Moved '+q+' pcs from "'+(source.internal||source.name)+'" to "'+(target.internal||target.name)+'" ✓');
    onMoved?.();
  };
  return <Modal title={'Move stock from "'+(source.internal||source.name)+'"…'} onClose={onClose} width={440}>
    <TC/>
    <div style={{fontSize:12,color:C.sub,marginBottom:12}}>Only the chosen quantity moves — both items stay, and no sale/purchase history is reassigned.</div>
    {!target
      ?<>
        <input value={pq} onChange={e=>setPq(e.target.value)} placeholder="Search item to move stock into…" autoFocus
          style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit',marginBottom:8,boxSizing:'border-box'}}/>
        {pq.length>0&&candidates.map(it=><div key={it.id} onClick={()=>setTargetId(it.id)} style={{padding:'8px 12px',borderRadius:7,cursor:'pointer',border:`1px solid ${C.bdr}`,marginBottom:6}}
          {...(mobile?{}:{onMouseEnter:e=>e.currentTarget.style.background='#F6F8FA',onMouseLeave:e=>e.currentTarget.style.background='transparent'})}>
          <div style={{fontWeight:600,color:C.txt,fontSize:13}}>{it.internal||it.name}</div>
          <div style={{fontSize:11,color:C.sub}}>{it.cat||'—'} · {it.stock||0} pcs in stock</div>
        </div>)}
      </>
      :<div style={{padding:'10px 14px',borderRadius:8,border:`1.5px solid ${C.blue}`,background:C.blueSoft,marginBottom:14}}>
        <div style={{fontSize:11,color:C.sub}}>Moving into</div>
        <div style={{fontWeight:700,color:C.txt,fontSize:14}}>{target.internal||target.name}</div>
        <div style={{fontSize:11,color:C.sub,marginTop:2}}>{target.stock||0} pcs → {(target.stock||0)+q} pcs after move</div>
        <button onClick={()=>setTargetId('')} style={{marginTop:4,background:'none',border:'none',color:C.blue,cursor:'pointer',fontSize:11,textDecoration:'underline',fontFamily:'inherit'}}>Change</button>
      </div>}
    {target&&<div style={{marginTop:4}}>
      <label style={{fontSize:11,color:C.sub,fontWeight:600}}>Quantity to move (of {source.stock||0} pcs)</label>
      <input type="number" value={qty} onChange={e=>setQty(e.target.value)}
        style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit',marginTop:4,boxSizing:'border-box'}}/>
    </div>}
    <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:10}}>
      <Btn onClick={onClose} variant="secondary">Cancel</Btn>
      <Btn onClick={doMove} disabled={!targetId||q<=0} variant="primary">Move</Btn>
    </div>
  </Modal>;
}

// ── PARTY MASTERS ─────────────────────────────────────────────────────────────
function PartyMasters(){
  const mobile=useIsMobile();
  const[parties,setParties]=useState(state.parties);
  const[search,setSearch]=useState('');
  const[routeF,setRouteF]=useState('');
  const[typeF,setTypeF]=useState('customer');
  const[modal,setModal]=useState(null);
  const[form,setForm]=useState({});
  const[editPhId,setEditPhId]=useState(null);
  const[editPhVal,setEditPhVal]=useState('');
  const[ledgerParty,setLedgerParty]=useState(null);
  const[mergeSrc,setMergeSrc]=useState(null);
  const{toast,TC}=useToast();
  const filtered=parties.filter(p=>p.type===typeF&&(!routeF||p.route===routeF)&&(!search||p.name.toLowerCase().includes(search.toLowerCase())||p.city?.toLowerCase().includes(search.toLowerCase())));
  const openEdit=p=>{setForm({...p});setModal('edit');};
  const doSave=()=>{
    if(!form.name?.trim())return toast('Name required','error');
    const p={...form,id:form.id||uid('PTY'),credit:Number(form.credit)||30,ob:Number(form.ob)||0};
    if(modal==='add')state.parties=[...state.parties,p]; else state.parties=state.parties.map(x=>x.id===p.id?p:x);
    state.phoneEdits[p.id]=p.phone||''; save.phoneEdits(); save.parties();
    setParties([...state.parties]); toast(modal==='add'?'Party added':'Updated'); setModal(null);
  };
  return <div>
    <TC/>
    <SectionHead title="Party Masters" sub={`${parties.filter(p=>p.type==='customer').length} customers`} action={<Btn onClick={()=>{setForm({type:'customer',credit:30,ob:0,active:true});setModal('add');}}>+ Add Party</Btn>}/>
    {ledgerParty&&<PartyLedger party={ledgerParty} onClose={()=>setLedgerParty(null)}/>}
    {mergeSrc&&<MergePartyModal source={mergeSrc} onClose={()=>setMergeSrc(null)} onMerged={()=>{setParties([...state.parties]);setMergeSrc(null);}}/>}
    {modal&&<Modal title={modal==='add'?'Add Party':'Edit Party'} onClose={()=>setModal(null)}>
      <Grid2>
        <div style={{gridColumn:'1/-1'}}><Inp label="Name" value={form.name||''} onChange={v=>setForm(f=>({...f,name:v}))} required/></div>
        <Sel label="Type" value={form.type||'customer'} onChange={v=>setForm(f=>({...f,type:v}))} options={[{value:'customer',label:'Customer'},{value:'supplier',label:'Supplier'},{value:'transporter',label:'Transporter'}]}/>
        <Sel label="Route" value={form.route||''} onChange={v=>setForm(f=>({...f,route:v}))} placeholder="Select route" options={[...ROUTES,'Other']}/>
        <Inp label="City" value={form.city||''} onChange={v=>setForm(f=>({...f,city:v}))}/>
        <Inp label="Phone" value={form.phone||''} onChange={v=>setForm(f=>({...f,phone:v}))} hint="WhatsApp number"/>
        <Inp label="Credit Days" value={form.credit||30} onChange={v=>setForm(f=>({...f,credit:Number(v)}))} type="number"/>
        <Inp label="Opening Balance ₹" value={form.ob||0} onChange={v=>setForm(f=>({...f,ob:Number(v)}))} type="number"/>
        <Inp label="GSTIN" value={form.gstin||''} onChange={v=>setForm(f=>({...f,gstin:v}))}/>
      </Grid2>
      <div style={{display:'flex',justifyContent:'flex-end',gap:10,marginTop:20}}>
        <Btn onClick={()=>setModal(null)} variant="secondary">Cancel</Btn>
        <Btn onClick={doSave}>{modal==='add'?'Add Party':'Save Changes'}</Btn>
      </div>
    </Modal>}
    <div style={{display:'flex',gap:8,marginBottom:12,flexWrap:'wrap'}}>
      <Chip v="Customers" active={typeF==='customer'} onClick={()=>setTypeF('customer')}/>
      <Chip v="Suppliers" active={typeF==='supplier'} onClick={()=>setTypeF('supplier')}/>
      <Chip v="Transporters" active={typeF==='transporter'} onClick={()=>setTypeF('transporter')}/>
      {typeF==='customer'&&ROUTES.map(r=><Chip key={r} v={r} active={routeF===r} onClick={()=>setRouteF(routeF===r?'':r)}/>)}
    </div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'8px 12px',marginBottom:12}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search parties…" style={{width:'100%',border:'none',outline:'none',fontSize:13,color:C.txt,fontFamily:'inherit',background:'transparent'}}/>
    </div>
    <div style={{background:mobile?'transparent':C.surf,borderRadius:8,border:mobile?'none':`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <CardTable mobile={mobile} rows={filtered.slice(0,100)} keyFn={(p,i)=>p.id||i}
        columns={[
          {key:'name',label:'Name',render:p=><span onClick={()=>setLedgerParty(p)} style={{fontWeight:600,color:C.blue,cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted'}}>{p.name}</span>},
          {key:'city',label:'City',render:p=><span style={{color:C.sub,fontSize:12}}>{p.city||'—'}</span>},
          {key:'route',label:'Route',render:p=><Badge v={p.route||'—'} color={p.route==='CG North'?'blue':p.route==='CG South'?'green':p.route==='Odisha'?'amber':'gray'}/>},
          {key:'od',label:(typeF==='supplier'||typeF==='transporter')?'Payable ₹':'OD ₹',align:'right',render:p=>{
            const od=(typeF==='supplier'||typeF==='transporter')?getSupplierBal(p.id):getCustomerBal(p.id);
            return <span style={{fontWeight:od!==0?700:400,color:od>0?C.red:od<0?C.green:C.muted}}>{od!==0?(od<0?'−':'')+fmt(Math.abs(od)):'—'}</span>;
          }},
          {key:'phone',label:'Phone',render:p=>
            editPhId===p.id
              ?<div style={{display:'flex',gap:4,alignItems:'center'}}>
                <input value={editPhVal} onChange={e=>setEditPhVal(e.target.value)} placeholder="10-digit" autoFocus autoComplete="off"
                  onKeyDown={e=>{if(e.key==='Enter'){updatePartyPhone(p.id,editPhVal);setParties([...state.parties]);setEditPhId(null);toast('Saved ✓');}if(e.key==='Escape')setEditPhId(null);}}
                  style={{width:130,padding:'4px 8px',borderRadius:6,border:`1.5px solid ${C.blue}`,fontSize:12,outline:'none',fontFamily:'inherit'}}/>
                <button onClick={()=>{updatePartyPhone(p.id,editPhVal);setParties([...state.parties]);setEditPhId(null);toast('Saved ✓');}} style={{padding:'4px 10px',borderRadius:6,background:C.blue,color:'#fff',border:'none',cursor:'pointer',fontSize:12,fontFamily:'inherit',fontWeight:600}}>✓</button>
                <button onClick={()=>setEditPhId(null)} style={{padding:'4px 8px',borderRadius:6,background:'none',color:C.muted,border:`1px solid ${C.bdr}`,cursor:'pointer',fontSize:12}}>✕</button>
              </div>
              :<button onClick={()=>{setEditPhId(p.id);setEditPhVal(p.phone||'');}} style={{padding:'4px 10px',borderRadius:6,border:`1px solid ${C.bdr}`,background:p.phone?C.greenSoft:C.amberSoft,color:p.phone?C.green:C.amber,fontSize:11,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>
                {p.phone?'📞 '+p.phone:'+ Add phone'}
              </button>},
          {key:'actions',label:'',cardLabel:'Actions',align:'right',render:p=>
            <div style={{display:'flex',gap:6,justifyContent:'flex-end',flexWrap:'wrap'}}>
              <Btn onClick={()=>setMergeSrc(p)} variant="ghost" size="sm">Merge</Btn>
              <Btn onClick={()=>openEdit(p)} variant="ghost" size="sm">Edit</Btn>
            </div>},
        ]}/>
    </div>
  </div>;
}

// ── SETTINGS ──────────────────────────────────────────────────────────────────
function SalesLog() {
  const[search,setSearch]=useState('');
  const[dateFrom,setDateFrom]=useState('');
  const[dateTo,setDateTo]=useState('');
  const[page,setPage]=useState(0);
  const[editSale,setEditSale]=useState(null);
  const[retSale,setRetSale]=useState(false);
  const[viewSale,setViewSale]=useState(null);
  const[ledgerParty,setLedgerParty]=useState(null);
  const[_fvSL,_setFvSL]=useState(0); const forceUpdate=()=>_setFvSL(v=>v+1);
  const PAGE=50;

  // Build from PARTY_HISTORY — no extra data needed
  const allSales = useMemo(()=>{
    const rows=[];
    for(const[party,hist] of Object.entries(PARTY_HISTORY)){
      for(const s of (hist.s||[])){
        for(const seg of splitPosNegItems(s[1],s[2]))rows.push({d:s[0],party,total:seg.total,items:seg.items,isReturn:seg.isReturn});
      }
    }
    // Also include new ERP sales
    for(const s of state.sales){
      rows.push({d:s.date,party:s.partyName,total:s.total,items:s.items?.map(x=>([x.customerName||x.itemName,x.qty,x.rate]))||[],isNew:true,isReturn:s.type==='return'});
    }
    return rows.sort((a,b)=>b.d.localeCompare(a.d));
  },[state.sales.length,_fvSL]);

  const filtered=useMemo(()=>allSales.filter(s=>
    (!search||s.party.toLowerCase().includes(search.toLowerCase())||s.items.some(x=>(x[0]||x).toLowerCase?.().includes(search.toLowerCase())))&&
    (!dateFrom||s.d>=dateFrom)&&(!dateTo||s.d<=dateTo)
  ),[allSales,search,dateFrom,dateTo]);

  const pageData=filtered.slice(page*PAGE,(page+1)*PAGE);
  const totalAmt=filtered.reduce((s,r)=>s+r.total,0);
  const totalPages=Math.ceil(filtered.length/PAGE);

  return <div>
    {editSale&&(editSale.type==='return'
      ?<EditReturnModal sale={editSale} onClose={()=>setEditSale(null)} onSaved={()=>forceUpdate(n=>n+1)}/>
      :<EditSaleModal sale={editSale} onClose={()=>setEditSale(null)} onSaved={()=>forceUpdate(n=>n+1)}/>)}
    {viewSale&&<InvoiceModal sale={viewSale} party={state.parties.find(p=>p.id===viewSale.partyId)||{name:viewSale.partyName,phone:''}} onClose={()=>setViewSale(null)} viewOnly/>}
    {retSale!==false&&<SalesReturnModal sale={retSale||null} onClose={()=>setRetSale(false)} onSaved={()=>forceUpdate(n=>n+1)}/>}
    {ledgerParty&&<PartyLedger party={ledgerParty} onClose={()=>setLedgerParty(null)}/>}
    <SectionHead title="Sales Log" sub={`${filtered.filter(s=>!s.isReturn).length} invoices · ${fmtL(filtered.filter(s=>!s.isReturn).reduce((s,r)=>s+r.total,0))} · ${filtered.filter(s=>s.isReturn).length>0?filtered.filter(s=>s.isReturn).length+' returns −'+fmtL(Math.abs(filtered.filter(s=>s.isReturn).reduce((s,r)=>s+r.total,0))):''}`} action={<Btn onClick={()=>setRetSale(null)} variant='danger' size='sm'>+ New Return</Btn>}/>
    <div style={{display:'flex',gap:10,marginBottom:14,flexWrap:'wrap'}}>
      <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="Search party or item…"
        style={{flex:1,minWidth:200,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}
        onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
      <input type="date" value={dateFrom} onChange={e=>{setDateFrom(e.target.value);setPage(0);}}
        style={{padding:'8px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
      <input type="date" value={dateTo} onChange={e=>{setDateTo(e.target.value);setPage(0);}}
        style={{padding:'8px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
      {(search||dateFrom||dateTo)&&<Btn onClick={()=>{setSearch('');setDateFrom('');setDateTo('');setPage(0);}} variant="secondary" size="sm">Clear</Btn>}
    </div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>
          {['Date','Party','Items','Amount',''].map((h,i)=><th key={i} style={{padding:'9px 14px',textAlign:i===3?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}
        </tr></thead>
        <tbody>
          {pageData.length===0&&<tr><td colSpan={4} style={{padding:'40px',textAlign:'center',color:C.muted}}>No entries match</td></tr>}
          {pageData.map((s,i)=>{
            const viewSaleForRow=()=>{
              if(s.isNew){
                const found=state.sales.find(x=>x.date===s.d&&x.partyName===s.party&&Math.abs(x.total-s.total)<1);
                if(found)setViewSale(found);
              }else{
                const p=state.parties.find(pp=>pp.name===s.party);
                setViewSale({id:null,date:s.d,partyId:p?.id,partyName:s.party,total:Math.abs(s.total),creditDays:p?.credit||30,
                  items:(s.items||[]).map(x=>Array.isArray(x)?{itemId:null,itemName:x[0],customerName:x[0],qty:x[1],rate:x[2],amount:x[1]*x[2]}:x)});
              }
            };
            return <tr key={i} onClick={viewSaleForRow} style={{borderBottom:`1px solid ${C.bdr}`,background:s.isReturn?C.redSoft:s.isNew?C.blueSoft:'transparent',cursor:'pointer'}}
            onMouseEnter={e=>e.currentTarget.style.background=s.isReturn?C.redSoft:s.isNew?C.blueSoft:'#F9FAFB'}
            onMouseLeave={e=>e.currentTarget.style.background=s.isReturn?C.redSoft:s.isNew?C.blueSoft:'transparent'}>
            <td style={{padding:'9px 14px',whiteSpace:'nowrap',fontVariantNumeric:'tabular-nums',color:C.sub}}>{fmtD(s.d)}{s.isNew&&<span style={{fontSize:10,color:C.blue,fontWeight:700,marginLeft:5}}>NEW</span>}</td>
            <td style={{padding:'9px 14px',fontWeight:600,color:C.txt,maxWidth:200,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
              <span onClick={e=>{e.stopPropagation();const p=state.parties.find(pp=>pp.name===s.party);if(p)setLedgerParty(p);}}
                style={{cursor:'pointer',color:C.blue,textDecoration:'underline',textDecorationStyle:'dotted'}}>{s.party}</span>
            </td>
            <td style={{padding:'9px 14px',fontSize:12,color:C.sub}}>
              {s.items.length>0?s.items.slice(0,2).map((x,j)=><span key={j} style={{display:'block'}}>{Array.isArray(x)?`${x[0]} ×${x[1]} @₹${x[2]}`:(x.customerName||x.itemName)}</span>):'—'}
              {s.items.length>2&&<span style={{color:C.muted,fontSize:11}}>+{s.items.length-2} more</span>}
            </td>
            <td style={{padding:'9px 14px',textAlign:'right',fontWeight:700,color:s.isReturn?C.red:C.blue,fontVariantNumeric:'tabular-nums'}}>{s.isReturn&&'−'}{fmt(Math.abs(s.total))}</td>
            <td style={{padding:'9px 14px',textAlign:'center'}}>
              <div style={{display:'flex',gap:5,justifyContent:'center'}}>
                <IconBtn title="View invoice" color={C.blue} bg={C.blueSoft}
                  onClick={e=>{e.stopPropagation();viewSaleForRow();}}>👁</IconBtn>
                {s.isNew&&<IconBtn title="Edit" color={C.sub} bg={C.surf}
                  onClick={e=>{e.stopPropagation();setEditSale(state.sales.find(x=>x.date===s.d&&x.partyName===s.party&&x.total===s.total)||null);}}>✎</IconBtn>}
                {!s.isReturn&&<IconBtn title="Record return" color={C.red} bg={C.redSoft}
                  onClick={e=>{e.stopPropagation();
                    if(s.isNew){const found=state.sales.find(x=>x.date===s.d&&x.partyName===s.party&&x.total===s.total);setRetSale(found||null);}
                    else setRetSale({partyName:s.party,partyId:state.parties.find(p=>p.name===s.party)?.id||'',date:s.d,total:s.total,items:s.items||[]});
                  }}>↩</IconBtn>}
                {s.isNew&&<IconBtn title="Delete" color={C.red} bg={C.redSoft}
                  onClick={e=>{e.stopPropagation();
                    if(!window.confirm('Delete this sale?'))return;
                    const found=state.sales.find(x=>x.date===s.d&&x.partyName===s.party&&x.total===s.total);
                    if(found){state.sales=state.sales.filter(x=>x.id!==found.id);save.sales();(found.items||[]).forEach(l=>adjustStock(l.itemId,l.qty));forceUpdate();}
                  }}>🗑</IconBtn>}
              </div>
            </td>
          </tr>;})}
        </tbody>
      </table>
    </div>
    {totalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:14}}>
      <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {page+1} of {totalPages}</span>
      <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}
  </div>;
}

// ── DELIVERIES ────────────────────────────────────────────────────────────────
function DeliveriesModule(){
  const[search,setSearch]=useState('');
  const[ledgerParty,setLedgerParty]=useState(null);
  const[_fv,_setFv]=useState(0); const forceUpdate=()=>_setFv(v=>v+1);
  const{toast,TC}=useToast();

  const pending=useMemo(()=>state.sales
    .filter(s=>s.type!=='return'&&!s.delivered)
    .sort((a,b)=>a.date.localeCompare(b.date)),
    [state.sales.length,_fv]);

  const filtered=useMemo(()=>pending.filter(s=>!search||s.partyName.toLowerCase().includes(search.toLowerCase())),[pending,search]);

  const setLR=(id,val)=>{
    state.sales=state.sales.map(s=>s.id===id?{...s,bilty:val}:s);
    save.sales();
  };
  const markDelivered=id=>{
    state.sales=state.sales.map(s=>s.id===id?{...s,delivered:true,deliveredDate:today()}:s);
    save.sales();
    forceUpdate();
    toast('Marked delivered ✓');
  };
  const sendConfirmation=sale=>{
    const party=state.parties.find(p=>p.id===sale.partyId);
    if(!party?.phone){toast('No phone on file for this party','error');return;}
    const itemsText=(sale.items||[]).map(l=>'• '+(l.customerName||l.itemName)+' x '+l.qty+' pcs').join('\n');
    const text=applyTemplate('delivery',{partyName:sale.partyName,items:itemsText,lr:sale.bilty?'LR: '+sale.bilty:''});
    window.open('https://wa.me/91'+party.phone+'?text='+encodeURIComponent(text),'_blank');
  };

  return <div>
    <TC/>
    {ledgerParty&&<PartyLedger party={ledgerParty} onClose={()=>setLedgerParty(null)}/>}
    <SectionHead title="Deliveries" sub={`${filtered.length} pending delivery confirmation`}/>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'8px 12px',marginBottom:14}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search party…" style={{width:'100%',border:'none',outline:'none',fontSize:13,color:C.txt,fontFamily:'inherit',background:'transparent'}}/>
    </div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>
          {['Date','Party','Items','Amount','LR Number','Actions'].map((h,i)=><th key={i} style={{padding:'9px 14px',textAlign:i===3?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}
        </tr></thead>
        <tbody>
          {filtered.length===0&&<tr><td colSpan={6} style={{padding:'40px',textAlign:'center',color:C.muted}}>No pending deliveries ✓</td></tr>}
          {filtered.map(s=>{
            const party=state.parties.find(p=>p.id===s.partyId);
            return <tr key={s.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
              <td style={{padding:'9px 14px',whiteSpace:'nowrap',color:C.sub}}>{fmtD(s.date)}</td>
              <td style={{padding:'9px 14px',fontWeight:600,color:C.txt,maxWidth:180,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
                <span onClick={()=>party&&setLedgerParty(party)} style={{cursor:'pointer',color:C.blue,textDecoration:'underline',textDecorationStyle:'dotted'}}>{s.partyName}</span>
              </td>
              <td style={{padding:'9px 14px',fontSize:12,color:C.sub}}>
                {(s.items||[]).slice(0,2).map((x,j)=><span key={j} style={{display:'block'}}>{(x.customerName||x.itemName)} ×{x.qty}</span>)}
                {(s.items||[]).length>2&&<span style={{color:C.muted,fontSize:11}}>+{s.items.length-2} more</span>}
              </td>
              <td style={{padding:'9px 14px',textAlign:'right',fontWeight:700,color:C.blue}}>{fmt(s.total)}</td>
              <td style={{padding:'9px 14px'}}>
                <input defaultValue={s.bilty||''} placeholder="LR number" onBlur={e=>setLR(s.id,e.target.value)}
                  style={{width:100,padding:'4px 8px',borderRadius:5,border:`1px solid ${s.bilty?C.bdr:C.amber}`,fontSize:12,outline:'none',fontFamily:'inherit'}}/>
              </td>
              <td style={{padding:'9px 14px',textAlign:'center'}}>
                <div style={{display:'flex',gap:6,justifyContent:'center'}}>
                  <IconBtn title={party?.phone?'Send WhatsApp confirmation':'No phone on file'} color="#25D366" bg="#F0FDF4" onClick={()=>sendConfirmation(s)}>📲</IconBtn>
                  <IconBtn title="Mark delivered" color={C.green} bg={C.greenSoft} onClick={()=>markDelivered(s.id)}>✓</IconBtn>
                </div>
              </td>
            </tr>;
          })}
        </tbody>
      </table>
    </div>
  </div>;
}

// ── PURCHASE LOG ──────────────────────────────────────────────────────────────
function PurchaseEntry({onSaved}){
  const[date,setDate]=useState(today());
  const[receivedDate,setReceivedDate]=useState(today());
  const[partyId,setPartyId]=useState('');
  const[lines,setLines]=useState([]);
  const[input,setInput]=useState('');
  const[si,setSi]=useState(0);
  const[step,setStep]=useState('item');
  const[addPartyName,setAddPartyName]=useState(null);
  const[draft,setDraft]=useState({itemId:'',iName:'',cName:'',sName:'',peti:1,qty:'',rate:'',freightPerPeti:''});
  const[pq,setPq]=useState('');
  const[pIdx,setPIdx]=useState(0);
  const[paidNow,setPaidNow]=useState(false);
  const[paidAccount,setPaidAccount]=useState('cash');
  const[transporterId,setTransporterId]=useState('');
  const[pqTr,setPqTr]=useState('');
  const[pIdxTr,setPIdxTr]=useState(0);
  const[addTransporterName,setAddTransporterName]=useState(null);
  const[bilty,setBilty]=useState('');
  const[freightPerPeti,setFreightPerPeti]=useState('');
  const[extraCharges,setExtraCharges]=useState([]);
  const[saving,setSaving]=useState(false);
  const{toast,TC}=useToast();
  const stepInputId='purch-step-input';

  // Named extra charges on the purchase itself (loading, packing, octroi, GST,
  // etc.) — owed to the SAME supplier, unlike freight which goes to a separate
  // transporter, so these roll straight into the purchase's own total below.
  const addExtraCharge=()=>setExtraCharges(p=>[...p,{id:uid('EC'),name:'',amount:''}]);
  const editExtraCharge=(i,k,v)=>setExtraCharges(p=>p.map((c,j)=>j!==i?c:{...c,[k]:v}));
  const removeExtraCharge=i=>setExtraCharges(p=>p.filter((_,j)=>j!==i));
  const extraChargesTotal=useMemo(()=>extraCharges.reduce((s,c)=>s+(Number(c.amount)||0),0),[extraCharges]);
  // Names typed before, across every past purchase, so "Loading"/"Octroi"/
  // "Packing" etc. can be picked instead of retyped every time.
  const extraChargeNames=useMemo(()=>{
    const seen=new Map();
    for(const p of state.purchases) for(const c of (p.extraCharges||[])) if(c.name) seen.set(c.name.toLowerCase(),c.name);
    return [...seen.values()];
  },[state.purchases.length]);

  useEffect(()=>{
    if(step!=='item'){setTimeout(()=>{const el=document.getElementById(stepInputId);if(el)el.focus();},30);}
  },[step,draft.iName]);

  const supps=useMemo(()=>state.parties.filter(p=>p.type==='supplier').sort((a,b)=>a.name.localeCompare(b.name)),[state.parties.length]);
  const sel=useMemo(()=>supps.find(p=>p.id===partyId),[supps,partyId]);
  const psugg=useMemo(()=>pq.length<1?[]:supps.filter(p=>p.name.toLowerCase().includes(pq.toLowerCase())||p.city?.toLowerCase().includes(pq.toLowerCase())).slice(0,8),[supps,pq]);
  const total=useMemo(()=>lines.reduce((s,l)=>s+(l.qty*l.rate),0),[lines]);

  const transporters=useMemo(()=>state.parties.filter(p=>p.type==='transporter').sort((a,b)=>a.name.localeCompare(b.name)),[state.parties.length]);
  const selTr=useMemo(()=>transporters.find(p=>p.id===transporterId),[transporters,transporterId]);
  const psuggTr=useMemo(()=>pqTr.length<1?[]:transporters.filter(p=>p.name.toLowerCase().includes(pqTr.toLowerCase())||p.city?.toLowerCase().includes(pqTr.toLowerCase())).slice(0,8),[transporters,pqTr]);
  const totalPetis=useMemo(()=>lines.reduce((s,l)=>s+(l.qty/(l.peti||1)),0),[lines]);
  const freightAmount=useMemo(()=>Math.round(lines.reduce((s,l)=>s+(l.qty/(l.peti||1))*(Number(l.freightPerPeti)||0),0)),[lines]);
  const grandTotal=total+extraChargesTotal+freightAmount;

  useEffect(()=>{
    if(sel){setTimeout(()=>{const el=document.getElementById('purch-item-search-input');if(el)el.focus();},50);}
  },[sel]);

  const isugg=useMemo(()=>{
    const q=input.trim().toLowerCase();if(!q||step!=='item')return[];
    const words=q.split(/[\s,]+/).filter(Boolean);
    return state.items.filter(it=>it.active!==false).map(it=>{
      const cn=(it.customer||it.internal||it.name||'').toLowerCase();
      const intn=(it.internal||it.name||'').toLowerCase();
      const hay=[cn,intn,(it.cat||'').toLowerCase(),(it.supplier||'').toLowerCase()].join(' ');
      let score=0;
      for(const w of words){if(cn.includes(w))score+=10;if(intn.includes(w))score+=8;if(hay.includes(w))score+=4;if(cn.startsWith(w))score+=10;}
      if(cn.includes(q))score+=15;
      return{it,score};
    }).filter(x=>x.score>0).sort((a,b)=>b.score-a.score).slice(0,10).map(x=>x.it);
  },[input,step]);

  const pickItem=it=>{
    const rate=sel?getBestPurchRate(sel.name,it.id,it.purch):it.purch;
    setDraft({itemId:it.id,iName:it.internal||it.name||'',cName:it.customer||it.internal||'',sName:it.supplier||it.internal||'',peti:it.peti,qty:'',rate,freightPerPeti:freightPerPeti});
    setStep('qty');setInput('');setSi(0);
  };

  const commit=()=>{
    const qty=parseQty(draft.qty,draft.peti);
    if(isNaN(qty)||qty<=0)return toast('Invalid qty','error');
    setLines(p=>[...p,{id:uid('L'),itemId:draft.itemId,itemName:draft.iName,customerName:draft.cName,supplierName:draft.sName,peti:draft.peti,qty,rate:Number(draft.rate)||0,freightPerPeti:Number(draft.freightPerPeti)||0}]);
    setDraft({itemId:'',iName:'',cName:'',sName:'',peti:1,qty:'',rate:'',freightPerPeti:''});setStep('item');setInput('');setSi(0);
    setTimeout(()=>document.getElementById('purch-item-search-input')?.focus(),30);
  };

  const onKey=e=>{
    if(step==='item'){
      if(e.key==='ArrowDown'){e.preventDefault();setSi(i=>Math.min(i+1,(isugg.length||1)-1));}
      if(e.key==='ArrowUp')  {e.preventDefault();setSi(i=>Math.max(i-1,0));}
      if(e.key==='Enter'&&isugg.length){e.preventDefault();pickItem(isugg[si]||isugg[0]);}
      else if(e.key==='Enter'&&!input.trim()){e.preventDefault();document.getElementById('purch-transport-search-input')?.focus();}
      if(e.key==='Escape'){setInput('');setSi(0);}
    }else if(step==='qty'){
      if(e.key==='Enter'){e.preventDefault();if(!draft.qty)return;const q=parseQty(draft.qty,draft.peti);if(isNaN(q))return toast('Invalid qty','error');setDraft(d=>({...d,qty:String(q)}));setStep('rate');}
      if(e.key==='Escape'){setStep('item');setInput('');}
    }else if(step==='rate'){
      if(e.key==='Enter'){e.preventDefault();commit();}
      if(e.key==='Escape'){setStep('qty');}
    }
  };

  const editLine=(i,k,v)=>setLines(p=>p.map((l,j)=>j!==i?l:{...l,[k]:Number(v)||0}));

  const doSave=()=>{
    if(!partyId)return toast('Select a supplier','error');
    if(!lines.length)return toast('Add at least one item','error');
    setSaving(true);
    const id=uid('PUR');
    const allItems=lines.map(l=>({itemId:l.itemId,itemName:l.itemName,customerName:l.customerName,supplierName:l.supplierName,qty:l.qty,rate:l.rate,amount:l.qty*l.rate,peti:l.peti||1,
      freightPerPeti:Number(l.freightPerPeti)||0,freightAmount:(l.qty/(l.peti||1))*(Number(l.freightPerPeti)||0)}));
    const extraChargesClean=extraCharges.filter(c=>c.name.trim()&&Number(c.amount)>0).map(c=>({id:c.id,name:c.name.trim(),amount:Number(c.amount)||0}));
    const extraChargesTotalSaved=extraChargesClean.reduce((s,c)=>s+c.amount,0);
    const supplierTotal=total+extraChargesTotalSaved;
    const purchase={id,date,bookedDate:date,receivedDate:receivedDate||null,partyId,partyName:sel?.name||'',items:allItems,total,bilty,
      extraCharges:extraChargesClean,extraChargesTotal:extraChargesTotalSaved,
      transporterId:selTr?.id||undefined,transporterName:selTr?.name||undefined,
      totalPetis,freightAmount,
      createdAt:new Date().toISOString()};
    state.purchases=[purchase,...state.purchases];save.purchases();
    if(receivedDate)lines.forEach(l=>adjustStock(l.itemId,l.qty));
    if(paidNow&&supplierTotal>0){
      const t={id:uid('TXN'),date,type:'payment',account:paidAccount,amount:supplierTotal,payee:sel?.name||'',partyId,partyName:sel?.name||'',category:'vendor',notes:'Purchase '+id,purchaseId:id,createdAt:new Date().toISOString()};
      state.txns=[t,...state.txns];save.txns();
      const sp={id:uid('SPMT'),txnId:t.id,date,partyId,partyName:sel?.name||'',amount:supplierTotal,account:paidAccount,notes:'Purchase '+id,purchaseId:id};
      state.supplierPayments=[sp,...state.supplierPayments];save.supplierPayments();
    }
    if(selTr&&freightAmount>0){
      const frt={id:uid('FRT'),date,transporterId:selTr.id,transporterName:selTr.name,purchaseId:id,
        totalPetis,ratePerPeti:totalPetis>0?freightAmount/totalPetis:0,amount:freightAmount,notes:'Freight for purchase '+id+(bilty?' · LR '+bilty:''),createdAt:new Date().toISOString()};
      state.freightCharges=[frt,...state.freightCharges];save.freightCharges();
    }
    setSaving(false);
    toast('Purchase recorded ✓ '+fmt(grandTotal));
    setPartyId('');setPq('');setLines([]);setStep('item');setInput('');setPaidNow(false);
    setTransporterId('');setPqTr('');setBilty('');setFreightPerPeti('');setExtraCharges([]);
    setDate(today());setReceivedDate(today());
    onSaved?.();
  };

  const SC={item:C.blue,qty:C.green,rate:C.amber};
  const SL={item:'Type anything — name, category, model…',qty:'Qty, "1 peti" or "1p" = '+draft.peti+'pc',rate:'Rate Rs/pc — Enter to add'};

  return <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:16,marginBottom:18}}>
    <TC/>
    {addPartyName!==null&&<QuickAddPartyModal defaultName={addPartyName} type="supplier" onClose={()=>setAddPartyName(null)} onAdded={p=>{setPartyId(p.id);setPq('');setAddPartyName(null);}}/>}
    <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:12}}>+ New Purchase</div>

    <div style={{display:'flex',gap:10,marginBottom:4,alignItems:'flex-start'}}>
      <div style={{flex:'0 0 155px'}}><Inp label="Booked Date" value={date} onChange={setDate} type="date" required/></div>
      <div style={{flex:'0 0 155px'}}><Inp label="Received Date" value={receivedDate} onChange={setReceivedDate} type="date" hint="Blank = not received yet"/></div>
    </div>
    {!receivedDate&&<div style={{fontSize:11,color:C.amber,fontWeight:600,marginBottom:10}}>⏳ Stock won't update until received</div>}

    {/* Supplier */}
    <div style={{marginBottom:10}}>
      <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:4}}>Supplier <span style={{color:C.red}}>*</span></label>
      {sel
        ?<div style={{display:'flex',alignItems:'center',gap:10,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
          <div style={{flex:1}}>
            <span style={{fontWeight:700,color:C.txt}}>{sel.name}</span>
            <span style={{fontSize:11,color:C.sub,marginLeft:8}}>{sel.city}</span>
          </div>
          <Btn onClick={()=>{setPartyId('');setPq('');}} variant="ghost" size="sm">Change</Btn>
        </div>
        :<div style={{position:'relative'}}>
          <input id="purch-supplier-search-input" autoFocus value={pq} autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
            onChange={e=>{setPq(e.target.value);setPIdx(0);}}
            onKeyDown={e=>{
              if(e.key==='ArrowDown'){e.preventDefault();setPIdx(i=>Math.min(i+1,(psugg.length||1)-1));}
              if(e.key==='ArrowUp')  {e.preventDefault();setPIdx(i=>Math.max(i-1,0));}
              if(e.key==='Enter'&&psugg.length){e.preventDefault();setPartyId(psugg[pIdx]?.id||psugg[0].id);setPq('');}
              else if(e.key==='Enter'&&pq.length>=2){e.preventDefault();setAddPartyName(pq);}
              if(e.key==='Escape'){setPq('');setPIdx(0);}
            }}
            placeholder="Start typing supplier name or city…"
            style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
            onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
          {(psugg.length>0||pq.length>=2)&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
            {psugg.map((p,pi)=><div key={p.id} onMouseDown={()=>{setPartyId(p.id);setPq('');}} onMouseEnter={()=>setPIdx(pi)}
              style={{padding:'9px 14px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`,background:pi===pIdx?C.blueSoft:'transparent'}}>
              <span style={{fontWeight:600,color:C.txt}}>{p.name}</span>
              <span style={{fontSize:11,color:C.sub,marginLeft:8}}>{p.city}</span>
            </div>)}
            {pq.length>=2&&psugg.length===0&&<div onMouseDown={()=>setAddPartyName(pq)}
              style={{padding:'9px 14px',cursor:'pointer',color:C.blue,fontWeight:600,fontSize:13,background:C.blueSoft}}>
              + Add "{pq}" as new supplier
            </div>}
          </div>}
        </div>}
    </div>

    {/* Item entry + lines */}
    <div style={{background:'#FAFBFC',borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',marginBottom:10}}>
      <div style={{padding:'10px 14px',background:step==='item'?C.blueSoft:step==='qty'?C.greenSoft:C.amberSoft,borderBottom:`1px solid ${C.bdr}`,position:'relative'}}>
        <div style={{display:'flex',gap:8,alignItems:'center'}}>
          {step!=='item'&&<span style={{fontSize:12,fontWeight:600,color:C.txt,whiteSpace:'nowrap',maxWidth:140,overflow:'hidden',textOverflow:'ellipsis'}}>
            {step==='rate'?`${draft.sName||draft.iName} × ${draft.qty}pc`:draft.sName||draft.iName}
          </span>}
          <input id={step==='item'?'purch-item-search-input':stepInputId} key={step+draft.iName}
            autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
            value={step==='item'?input:step==='qty'?draft.qty:String(draft.rate||'')}
            onChange={e=>{const v=e.target.value;if(step==='item'){setInput(v);setSi(0);}else if(step==='qty')setDraft(d=>({...d,qty:v}));else setDraft(d=>({...d,rate:v}));}}
            onKeyDown={onKey} placeholder={SL[step]}
            style={{flex:1,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${SC[step]}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',fontWeight:500}}/>
          {step!=='item'&&<><Btn onClick={commit} size="sm" variant="success">Add</Btn><Btn onClick={()=>{setStep('item');setInput('');}} size="sm" variant="ghost">✕</Btn></>}
        </div>
        {step==='item'&&isugg.length>0&&<div style={{position:'absolute',top:'calc(100% - 2px)',left:14,right:14,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden',maxHeight:180,display:'flex',flexDirection:'column'}}>
          <div style={{overflowY:'auto',flex:1}}>
            {isugg.map((it,i)=>{
              const r=sel?getBestPurchRate(sel.name,it.id,it.purch):it.purch;
              return <div key={it.id} id={'psugg-item-'+i} onMouseDown={()=>{setSi(i);pickItem(it);}} onMouseEnter={()=>setSi(i)}
                style={{display:'flex',alignItems:'center',padding:'8px 14px',cursor:'pointer',background:i===si?C.blueSoft:'transparent',borderBottom:i<isugg.length-1?`1px solid ${C.bdr}`:'none'}}>
                <div style={{flex:1}}>
                  <span style={{fontWeight:600,color:C.txt,fontSize:13}}>{it.supplier||it.internal}</span>
                  {it.internal&&it.internal!==(it.supplier||it.internal)&&<span style={{fontSize:10,color:C.muted,marginLeft:6}}>{it.internal}</span>}
                  <span style={{fontSize:11,color:C.sub,marginLeft:6}}>{it.cat} · {it.peti}pc/peti · stock {it.stock||0}</span>
                </div>
                <div style={{textAlign:'right',flexShrink:0}}>
                  <span style={{fontSize:13,fontWeight:700,color:r!==it.purch?C.amber:C.blue}}>₹{r}/pc</span>
                  {r!==it.purch&&<div style={{fontSize:10,color:C.muted}}>std ₹{it.purch}</div>}
                </div>
              </div>;
            })}
          </div>
          <div style={{padding:'4px 14px',fontSize:10,color:C.muted,background:'#F6F8FA',flexShrink:0}}>↑↓ · Enter select · Esc cancel</div>
        </div>}
        {step==='qty'&&draft.qty&&!isNaN(parseQty(draft.qty,draft.peti))&&<div style={{marginTop:4,fontSize:11,color:C.green}}>→ {parseQty(draft.qty,draft.peti)} pcs</div>}
      </div>

      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>
          {['#','Item','Qty','Rate ₹','Freight ₹/peti','Amount',''].map((h,i)=><th key={i} style={{padding:'7px 12px',textAlign:i>1&&i<6?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}
        </tr></thead>
        <tbody>
          {!lines.length&&<tr><td colSpan={7} style={{padding:'20px',textAlign:'center',color:C.muted,fontSize:12}}>{sel?'Type item or model above':'Select a supplier first'}</td></tr>}
          {lines.map((l,i)=><tr key={l.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
            <td style={{padding:'7px 12px',color:C.muted,fontSize:11}}>{i+1}</td>
            <td style={{padding:'7px 12px'}}>
              <div style={{fontWeight:600,color:C.txt}}>{l.supplierName||l.itemName}</div>
              {l.itemName&&l.itemName!==(l.supplierName||l.itemName)&&<div style={{fontSize:10,color:C.muted}}>{l.itemName}</div>}
            </td>
            <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.qty} onChange={e=>editLine(i,'qty',e.target.value)} type="number" style={{width:60,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
            <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.rate} onChange={e=>editLine(i,'rate',e.target.value)} type="number" style={{width:75,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
            <td style={{padding:'7px 12px',textAlign:'right'}}>
              <input value={l.freightPerPeti||''} onChange={e=>editLine(i,'freightPerPeti',e.target.value)} type="number" placeholder="0" style={{width:70,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/>
              {Number(l.freightPerPeti)>0&&<div style={{fontSize:10,color:C.muted,marginTop:2}}>{fmt(Math.round((l.qty/(l.peti||1))*Number(l.freightPerPeti)))}</div>}
            </td>
            <td style={{padding:'7px 12px',textAlign:'right',fontWeight:700,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(l.qty*l.rate)}</td>
            <td style={{padding:'7px 12px',textAlign:'center'}}><button onClick={()=>setLines(p=>p.filter((_,j)=>j!==i))} style={{background:'none',border:'none',cursor:'pointer',color:C.muted,fontSize:15}}>×</button></td>
          </tr>)}
        </tbody>
        {lines.length>0&&<tfoot><tr style={{background:'#F6F8FA'}}>
          <td colSpan={5} style={{padding:'9px 12px',fontWeight:700,textAlign:'right',color:C.txt}}>TOTAL</td>
          <td style={{padding:'9px 12px',textAlign:'right',fontSize:16,fontWeight:800,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(total)}</td>
          <td/>
        </tr></tfoot>}
      </table>
    </div>

    {/* Transporter + Freight */}
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'12px 16px',marginBottom:10}}>
      <Grid2>
        <div style={{position:'relative'}}>
          <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:4}}>Transporter</label>
          {addTransporterName!==null&&<QuickAddPartyModal defaultName={addTransporterName} type="transporter" onClose={()=>setAddTransporterName(null)} onAdded={p=>{setTransporterId(p.id);setPqTr('');setAddTransporterName(null);}}/>}
          {selTr
            ?<div style={{display:'flex',alignItems:'center',gap:10,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
              <div style={{flex:1}}><span style={{fontWeight:700,color:C.txt}}>{selTr.name}</span></div>
              <Btn onClick={()=>{setTransporterId('');setPqTr('');}} variant="ghost" size="sm">Change</Btn>
            </div>
            :<div style={{position:'relative'}}>
              <input id="purch-transport-search-input" value={pqTr} autoComplete="off" autoCorrect="off"
                onChange={e=>{setPqTr(e.target.value);setPIdxTr(0);}}
                onKeyDown={e=>{
                  if(e.key==='ArrowDown'){e.preventDefault();setPIdxTr(i=>Math.min(i+1,(psuggTr.length||1)-1));}
                  if(e.key==='ArrowUp')  {e.preventDefault();setPIdxTr(i=>Math.max(i-1,0));}
                  if(e.key==='Enter'&&psuggTr.length){e.preventDefault();setTransporterId(psuggTr[pIdxTr]?.id||psuggTr[0].id);setPqTr('');}
                  else if(e.key==='Enter'&&pqTr.length>=2){e.preventDefault();setAddTransporterName(pqTr);}
                }}
                placeholder="Type transporter name…"
                style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
                onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
              {(psuggTr.length>0||pqTr.length>=2)&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
                {psuggTr.map((p,pi)=><div key={p.id} onMouseDown={()=>{setTransporterId(p.id);setPqTr('');}} onMouseEnter={()=>setPIdxTr(pi)}
                  style={{padding:'9px 14px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`,background:pi===pIdxTr?C.blueSoft:'transparent'}}>
                  <span style={{fontWeight:600,color:C.txt}}>{p.name}</span>
                </div>)}
                {pqTr.length>=2&&psuggTr.length===0&&<div onMouseDown={()=>setAddTransporterName(pqTr)}
                  style={{padding:'9px 14px',cursor:'pointer',color:C.blue,fontWeight:600,fontSize:13,background:C.blueSoft}}>
                  + Add "{pqTr}" as new transporter
                </div>}
              </div>}
            </div>}
        </div>
        <Inp label="Bilty / LR No." value={bilty} onChange={setBilty} placeholder="LR number" autoComplete="off"/>
      </Grid2>
      <div style={{marginTop:10,display:'flex',alignItems:'flex-end',gap:10,flexWrap:'wrap'}}>
        <div style={{flex:'0 0 200px'}}><Inp label="Default Freight ₹/peti" value={freightPerPeti} onChange={setFreightPerPeti} type="number" placeholder="0" hint="Applied to items added from now on — each line's rate stays editable"/></div>
        {freightAmount>0&&<span style={{fontSize:12,color:C.sub,paddingBottom:8}}>
          {totalPetis.toFixed(2)} petis total · freight <b style={{color:C.txt}}>{fmt(Math.round(freightAmount))}</b>
        </span>}
      </div>
    </div>

    {/* Extra charges */}
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'12px 16px',marginBottom:10}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:extraCharges.length?8:0}}>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>Extra Charges <span style={{fontWeight:500,color:C.muted,textTransform:'none'}}>(loading, octroi, packing…)</span></label>
        <Btn onClick={addExtraCharge} variant="ghost" size="sm">+ Add Charge</Btn>
      </div>
      {extraCharges.map((c,i)=><div key={c.id} style={{display:'flex',gap:8,alignItems:'center',marginBottom:6,position:'relative'}}>
        <input value={c.name} onChange={e=>editExtraCharge(i,'name',e.target.value)} list="extra-charge-names" placeholder="Charge name — e.g. Loading"
          style={{flex:1,padding:'7px 10px',borderRadius:7,border:`1px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit'}}/>
        <input value={c.amount} onChange={e=>editExtraCharge(i,'amount',e.target.value)} type="number" placeholder="Amount"
          style={{width:110,padding:'7px 10px',borderRadius:7,border:`1px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',textAlign:'right'}}/>
        <button onClick={()=>removeExtraCharge(i)} style={{background:'none',border:'none',cursor:'pointer',color:C.muted,fontSize:16}}>×</button>
      </div>)}
      <datalist id="extra-charge-names">{extraChargeNames.map(n=><option key={n} value={n}/>)}</datalist>
      {extraChargesTotal>0&&<div style={{fontSize:12,color:C.sub,textAlign:'right'}}>Extra charges total: <b style={{color:C.txt}}>{fmt(extraChargesTotal)}</b></div>}
    </div>

    {/* Paid now */}
    <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:14,flexWrap:'wrap'}}>
      <label style={{display:'flex',alignItems:'center',gap:6,fontSize:13,color:C.txt,cursor:'pointer',fontWeight:600}}>
        <input type="checkbox" checked={paidNow} onChange={e=>setPaidNow(e.target.checked)}/>Paid now
      </label>
      {paidNow&&<select value={paidAccount} onChange={e=>setPaidAccount(e.target.value)} style={{padding:'6px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:12,fontFamily:'inherit',color:C.txt}}>
        {ACCOUNTS.map(a=><option key={a.id} value={a.id}>{a.icon} {a.short}</option>)}
      </select>}
      {paidNow&&<span style={{fontSize:11,color:C.sub}}>Debits {fmt(total+extraChargesTotal)} on save, settles the supplier payable immediately{freightAmount>0?' (freight is tracked separately against the transporter)':''}</span>}
    </div>

    {(freightAmount>0||extraChargesTotal>0)&&<div style={{fontSize:12,color:C.sub,marginBottom:10,textAlign:'right'}}>Goods {fmt(total)}{extraChargesTotal>0?' + Charges '+fmt(extraChargesTotal):''}{freightAmount>0?' + Freight '+fmt(freightAmount):''} = <b style={{color:C.txt}}>{fmt(grandTotal)}</b></div>}

    <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
      <Btn onClick={()=>{setPartyId('');setPq('');setLines([]);setStep('item');setInput('');setPaidNow(false);setTransporterId('');setPqTr('');setBilty('');setFreightPerPeti('');setExtraCharges([]);setDate(today());setReceivedDate(today());}} variant="secondary" size="sm">Clear</Btn>
      <Btn onClick={doSave} disabled={saving||!partyId||!lines.length} size="md" icon="💾">{saving?'Saving…':'Save Purchase · '+fmt(grandTotal)}</Btn>
    </div>
  </div>;
}

// ── VIEW PURCHASE ─────────────────────────────────────────────────────────────
function PurchaseViewModal({purchase,onClose}){
  const items=purchase.items||[];
  const freightTotal=purchase.freightAmount||0;
  const extraChargesTotal=purchase.extraChargesTotal||0;
  const grandTotal=(purchase.total||0)+extraChargesTotal+freightTotal;
  return <Modal title={'Purchase — '+(purchase.partyName||'?')} onClose={onClose} width={600}>
    <div style={{display:'flex',gap:16,flexWrap:'wrap',marginBottom:14,fontSize:12,color:C.sub}}>
      <div><b style={{color:C.txt}}>Booked:</b> {fmtD(purchase.bookedDate||purchase.date)}</div>
      <div><b style={{color:C.txt}}>Received:</b> {purchase.receivedDate?fmtD(purchase.receivedDate):<span style={{color:C.amber}}>Pending</span>}</div>
      {purchase.transporterName&&<div><b style={{color:C.txt}}>Transporter:</b> {purchase.transporterName}</div>}
      {purchase.bilty&&<div><b style={{color:C.txt}}>LR:</b> {purchase.bilty}</div>}
    </div>
    <table style={{width:'100%',borderCollapse:'collapse',fontSize:13,marginBottom:14}}>
      <thead><tr style={{background:'#F6F8FA'}}>
        {['Item','Qty','Rate ₹','Freight ₹/p','Amount'].map((h,i)=><th key={i} style={{padding:'7px 12px',textAlign:i>0?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}
      </tr></thead>
      <tbody>
        {items.map((l,i)=><tr key={i} style={{borderBottom:`1px solid ${C.bdr}`}}>
          <td style={{padding:'7px 12px',fontWeight:600,color:C.txt}}>{l.supplierName||l.customerName||l.itemName}</td>
          <td style={{padding:'7px 12px',textAlign:'right'}}>{l.qty}</td>
          <td style={{padding:'7px 12px',textAlign:'right'}}>{fmt(l.rate)}</td>
          <td style={{padding:'7px 12px',textAlign:'right',color:C.sub}}>{l.freightPerPeti>0?fmt(l.freightPerPeti):'—'}</td>
          <td style={{padding:'7px 12px',textAlign:'right',fontWeight:700,color:C.blue}}>{fmt(l.qty*l.rate)}</td>
        </tr>)}
      </tbody>
      <tfoot>
        <tr style={{background:'#F6F8FA'}}><td colSpan={4} style={{padding:'9px 12px',fontWeight:700,textAlign:'right',color:C.txt}}>Goods Total</td><td style={{padding:'9px 12px',textAlign:'right',fontWeight:700,color:C.blue}}>{fmt(purchase.total)}</td></tr>
        {(purchase.extraCharges||[]).map(c=><tr key={c.id} style={{background:'#F6F8FA'}}><td colSpan={4} style={{padding:'9px 12px',fontWeight:600,textAlign:'right',color:C.sub}}>{c.name}</td><td style={{padding:'9px 12px',textAlign:'right',fontWeight:700,color:C.purple}}>{fmt(c.amount)}</td></tr>)}
        {freightTotal>0&&<tr style={{background:'#F6F8FA'}}><td colSpan={4} style={{padding:'9px 12px',fontWeight:700,textAlign:'right',color:C.txt}}>Freight</td><td style={{padding:'9px 12px',textAlign:'right',fontWeight:700,color:C.amber}}>{fmt(freightTotal)}</td></tr>}
        {(freightTotal>0||extraChargesTotal>0)&&<tr style={{background:'#F6F8FA'}}><td colSpan={4} style={{padding:'9px 12px',fontWeight:800,textAlign:'right',color:C.txt}}>Grand Total</td><td style={{padding:'9px 12px',textAlign:'right',fontSize:15,fontWeight:800,color:C.blue}}>{fmt(grandTotal)}</td></tr>}
      </tfoot>
    </table>
    <div style={{display:'flex',justifyContent:'flex-end'}}>
      <Btn onClick={onClose} variant="secondary">Close</Btn>
    </div>
  </Modal>;
}

// ── EDIT PURCHASE ──────────────────────────────────────────────────────────────
function EditPurchaseModal({purchase,onClose,onSaved}){
  const[bookedDate,setBookedDate]=useState(purchase.bookedDate||purchase.date||today());
  const[receivedDate,setReceivedDate]=useState(purchase.receivedDate||'');
  const[partyId,setPartyId]=useState(purchase.partyId||'');
  const[pq,setPq]=useState('');
  const[transporterId,setTransporterId]=useState(purchase.transporterId||'');
  const[pqTr,setPqTr]=useState('');
  const[bilty,setBilty]=useState(purchase.bilty||'');
  const[lines,setLines]=useState((purchase.items||[]).map((it,i)=>({id:'EL'+i,itemId:it.itemId||'',itemName:it.itemName||'',supplierName:it.supplierName||it.itemName||'',customerName:it.customerName||it.itemName||'',qty:it.qty||0,rate:it.rate||0,freightPerPeti:it.freightPerPeti||0,peti:it.peti||1})));
  const[extraCharges,setExtraCharges]=useState((purchase.extraCharges||[]).map(c=>({...c})));
  const addExtraCharge=()=>setExtraCharges(p=>[...p,{id:uid('EC'),name:'',amount:''}]);
  const editExtraCharge=(i,k,v)=>setExtraCharges(p=>p.map((c,j)=>j!==i?c:{...c,[k]:v}));
  const removeExtraCharge=i=>setExtraCharges(p=>p.filter((_,j)=>j!==i));
  const extraChargesTotal=useMemo(()=>extraCharges.reduce((s,c)=>s+(Number(c.amount)||0),0),[extraCharges]);
  const extraChargeNames=useMemo(()=>{
    const seen=new Map();
    for(const p of state.purchases) for(const c of (p.extraCharges||[])) if(c.name) seen.set(c.name.toLowerCase(),c.name);
    return [...seen.values()];
  },[state.purchases.length]);
  const[step,setStep]=useState('item');
  const[input,setInput]=useState('');
  const[si,setSi]=useState(0);
  const[draft,setDraft]=useState({itemId:'',iName:'',sName:'',peti:1,qty:'',rate:'',freightPerPeti:''});
  const{toast,TC}=useToast();
  const supps=useMemo(()=>state.parties.filter(p=>p.type==='supplier'),[state.parties.length]);
  const selParty=useMemo(()=>supps.find(p=>p.id===partyId),[supps,partyId]);
  const psugg=useMemo(()=>pq.length<1?[]:supps.filter(p=>p.name.toLowerCase().includes(pq.toLowerCase())).slice(0,6),[supps,pq]);
  const transporters=useMemo(()=>state.parties.filter(p=>p.type==='transporter'),[state.parties.length]);
  const selTr=useMemo(()=>transporters.find(p=>p.id===transporterId),[transporters,transporterId]);
  const psuggTr=useMemo(()=>pqTr.length<1?[]:transporters.filter(p=>p.name.toLowerCase().includes(pqTr.toLowerCase())).slice(0,6),[transporters,pqTr]);
  const total=useMemo(()=>lines.reduce((s,l)=>s+(Number(l.qty)||0)*(Number(l.rate)||0),0),[lines]);
  const freightAmount=useMemo(()=>Math.round(lines.reduce((s,l)=>s+((Number(l.qty)||0)/(l.peti||1))*(Number(l.freightPerPeti)||0),0)),[lines]);
  const totalPetis=useMemo(()=>lines.reduce((s,l)=>s+((Number(l.qty)||0)/(l.peti||1)),0),[lines]);
  useEffect(()=>{if(step!=='item'){setTimeout(()=>{const el=document.getElementById('edit-purch-step-input');if(el)el.focus();},30);}else{setTimeout(()=>{const el=document.getElementById('edit-purch-item-search');if(el)el.focus();},30);}},[step,draft.iName]);
  const isugg=useMemo(()=>{const q=input.trim().toLowerCase();if(!q||step!=='item')return[];const words=q.split(/[\s,]+/).filter(Boolean);return state.items.filter(it=>it.active!==false).map(it=>{const sn=(it.supplier||it.internal||'').toLowerCase();const hay=[sn,(it.internal||'').toLowerCase(),(it.cat||'').toLowerCase()].join(' ');let sc=0;for(const w of words){if(sn.includes(w))sc+=10;if(hay.includes(w))sc+=4;if(sn.startsWith(w))sc+=10;}if(sn.includes(q))sc+=15;return{it,sc};}).filter(x=>x.sc>0).sort((a,b)=>b.sc-a.sc).slice(0,8).map(x=>x.it);},[input,step]);
  const pickItem=it=>{const rate=selParty?getBestPurchRate(selParty.name,it.id,it.purch):it.purch;setDraft({itemId:it.id,iName:it.internal||it.name||'',sName:it.supplier||it.internal||'',peti:it.peti,qty:'',rate,freightPerPeti:''});setStep('qty');setInput('');setSi(0);};
  const commit=()=>{const qty=parseQty(draft.qty,draft.peti);if(isNaN(qty)||qty<=0)return toast('Invalid qty','error');setLines(p=>[...p,{id:uid('EL'),itemId:draft.itemId,itemName:draft.iName,supplierName:draft.sName,customerName:draft.sName,qty,rate:Number(draft.rate)||0,freightPerPeti:Number(draft.freightPerPeti)||0,peti:draft.peti}]);setDraft({itemId:'',iName:'',sName:'',peti:1,qty:'',rate:'',freightPerPeti:''});setStep('item');setInput('');setSi(0);setTimeout(()=>document.getElementById('edit-purch-item-search')?.focus(),30);};
  const onKey=e=>{if(step==='item'){if(e.key==='ArrowDown'){e.preventDefault();setSi(i=>Math.min(i+1,(isugg.length||1)-1));}if(e.key==='ArrowUp'){e.preventDefault();setSi(i=>Math.max(i-1,0));}if(e.key==='Enter'&&isugg.length){e.preventDefault();pickItem(isugg[si]||isugg[0]);}if(e.key==='Escape'){setInput('');setSi(0);}}else if(step==='qty'){if(e.key==='Enter'){e.preventDefault();if(!draft.qty)return;const q=parseQty(draft.qty,draft.peti);if(isNaN(q))return toast('Invalid qty','error');setDraft(d=>({...d,qty:String(q)}));setStep('rate');}if(e.key==='Escape'){setStep('item');setInput('');}}else if(step==='rate'){if(e.key==='Enter'){e.preventDefault();commit();}if(e.key==='Escape'){setStep('qty');}}};
  const editLine=(i,k,v)=>setLines(p=>p.map((l,j)=>j!==i?l:{...l,[k]:Number(v)||0}));
  const doSave=()=>{
    if(!lines.length)return toast('Add at least one item','error');
    const allItems=lines.map(l=>({itemId:l.itemId,itemName:l.itemName,customerName:l.customerName,supplierName:l.supplierName,qty:l.qty,rate:l.rate,amount:l.qty*l.rate,peti:l.peti||1,freightPerPeti:Number(l.freightPerPeti)||0,freightAmount:(l.qty/(l.peti||1))*(Number(l.freightPerPeti)||0)}));
    const extraChargesClean=extraCharges.filter(c=>c.name.trim()&&Number(c.amount)>0).map(c=>({id:c.id,name:c.name.trim(),amount:Number(c.amount)||0}));
    const extraChargesTotalSaved=extraChargesClean.reduce((s,c)=>s+c.amount,0);
    const updated={...purchase,bookedDate,date:bookedDate,receivedDate:receivedDate||null,partyId,partyName:selParty?.name||purchase.partyName,
      transporterId:selTr?.id||undefined,transporterName:selTr?.name||undefined,bilty,
      extraCharges:extraChargesClean,extraChargesTotal:extraChargesTotalSaved,
      items:allItems,total,freightAmount,totalPetis,updatedAt:new Date().toISOString()};
    state.purchases=state.purchases.map(p=>p.id===purchase.id?updated:p);save.purchases();

    const wasReceived=!!purchase.receivedDate, willBeReceived=!!receivedDate;
    if(wasReceived&&willBeReceived){
      const oldQty={},newQty={};
      (purchase.items||[]).forEach(l=>{oldQty[l.itemId]=(oldQty[l.itemId]||0)+l.qty;});
      lines.forEach(l=>{newQty[l.itemId]=(newQty[l.itemId]||0)+l.qty;});
      const ids=new Set([...Object.keys(oldQty),...Object.keys(newQty)]);
      ids.forEach(id=>{const delta=(newQty[id]||0)-(oldQty[id]||0); if(delta)adjustStock(id,delta);});
    }else if(!wasReceived&&willBeReceived){
      lines.forEach(l=>adjustStock(l.itemId,l.qty));
    }else if(wasReceived&&!willBeReceived){
      (purchase.items||[]).forEach(l=>adjustStock(l.itemId,-l.qty));
    }

    // Freight charge record against the transporter is derived state — drop the
    // old one and recreate from scratch rather than trying to patch it in place.
    state.freightCharges=state.freightCharges.filter(f=>f.purchaseId!==purchase.id);
    if(selTr&&freightAmount>0){
      const frt={id:uid('FRT'),date:bookedDate,transporterId:selTr.id,transporterName:selTr.name,purchaseId:purchase.id,
        totalPetis,ratePerPeti:totalPetis>0?freightAmount/totalPetis:0,amount:freightAmount,notes:'Freight for purchase '+purchase.id+(bilty?' · LR '+bilty:''),createdAt:new Date().toISOString()};
      state.freightCharges=[frt,...state.freightCharges];
    }
    save.freightCharges();

    toast('Purchase updated ✓');onSaved?.();onClose();
  };
  const SC={item:C.blue,qty:C.green,rate:C.amber};const SL={item:'Type item name to add…',qty:'Qty or "1 peti"',rate:'Rate ₹/pc'};
  return <Modal title={'Edit Purchase — '+(purchase.partyName||'?')} onClose={onClose} width={700}>
    <TC/>
    <Grid2>
      <Inp label="Booked Date" value={bookedDate} onChange={setBookedDate} type="date" required/>
      <Inp label="Received Date" value={receivedDate} onChange={setReceivedDate} type="date" hint="Blank = not received yet"/>
      <div style={{gridColumn:'1/-1'}}>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>Supplier</label>
        {selParty
          ?<div style={{display:'flex',alignItems:'center',gap:10,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
            <span style={{fontWeight:700,flex:1,color:C.txt}}>{selParty.name}</span>
            <Btn onClick={()=>{setPartyId('');setPq('');}} variant="ghost" size="sm">Change</Btn>
          </div>
          :<div style={{position:'relative'}}>
            <input value={pq} autoComplete="off" onChange={e=>setPq(e.target.value)} placeholder="Type supplier name…"
              style={{width:'100%',padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
              onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
            {psugg.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 8px 8px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
              {psugg.map(p=><div key={p.id} onMouseDown={()=>{setPartyId(p.id);setPq('');}} style={{padding:'9px 14px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`}} onMouseEnter={e=>e.currentTarget.style.background='#F9FAFB'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                <span style={{fontWeight:600,color:C.txt}}>{p.name}</span></div>)}
            </div>}
          </div>}
      </div>
      <div style={{position:'relative'}}>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>Transporter</label>
        {selTr
          ?<div style={{display:'flex',alignItems:'center',gap:8,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
            <span style={{fontWeight:700,flex:1,color:C.txt,fontSize:13}}>{selTr.name}</span>
            <Btn onClick={()=>{setTransporterId('');setPqTr('');}} variant="ghost" size="sm">Change</Btn>
          </div>
          :<div style={{position:'relative'}}>
            <input value={pqTr} autoComplete="off" onChange={e=>setPqTr(e.target.value)} placeholder="Type transporter…"
              style={{width:'100%',padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
              onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
            {psuggTr.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 8px 8px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
              {psuggTr.map(p=><div key={p.id} onMouseDown={()=>{setTransporterId(p.id);setPqTr('');}} style={{padding:'9px 14px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`}}><span style={{fontWeight:600,color:C.txt}}>{p.name}</span></div>)}
            </div>}
          </div>}
      </div>
      <Inp label="Bilty / LR" value={bilty} onChange={setBilty} placeholder="LR number" autoComplete="off"/>
    </Grid2>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',margin:'14px 0'}}>
      <div style={{padding:'10px 14px',background:SC[step]+'18',borderBottom:`1px solid ${C.bdr}`,position:'relative'}}>
        <div style={{display:'flex',gap:8,alignItems:'center'}}>
          {step!=='item'&&<span style={{fontSize:12,fontWeight:600,color:C.txt,whiteSpace:'nowrap',maxWidth:130,overflow:'hidden',textOverflow:'ellipsis'}}>{step==='rate'?`${draft.sName} × ${draft.qty}pc`:draft.sName}</span>}
          <input id={step==='item'?'edit-purch-item-search':'edit-purch-step-input'} key={step+draft.iName} autoComplete="off"
            value={step==='item'?input:step==='qty'?draft.qty:String(draft.rate||'')}
            onChange={e=>{const v=e.target.value;if(step==='item'){setInput(v);setSi(0);}else if(step==='qty')setDraft(d=>({...d,qty:v}));else setDraft(d=>({...d,rate:v}));}}
            onKeyDown={onKey} placeholder={SL[step]}
            style={{flex:1,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${SC[step]}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit'}}/>
          {step!=='item'&&<><Btn onClick={commit} size="sm" variant="success">Add ↵</Btn><Btn onClick={()=>{setStep('item');setInput('');}} size="sm" variant="ghost">✕</Btn></>}
        </div>
        {step==='item'&&isugg.length>0&&<div style={{position:'absolute',top:'calc(100% - 2px)',left:14,right:14,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 9px 9px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden',maxHeight:180,display:'flex',flexDirection:'column'}}>
          <div style={{overflowY:'auto',flex:1}}>
            {isugg.map((it,i)=><div key={it.id} id={'esugg-purch-'+i} onMouseDown={()=>{setSi(i);pickItem(it);}} onMouseEnter={()=>setSi(i)}
              style={{display:'flex',alignItems:'center',padding:'8px 14px',cursor:'pointer',background:i===si?C.blueSoft:'transparent',borderBottom:i<isugg.length-1?`1px solid ${C.bdr}`:'none'}}>
              <div style={{flex:1}}><span style={{fontWeight:600,color:C.txt,fontSize:13}}>{it.supplier||it.internal}</span><span style={{fontSize:11,color:C.sub,marginLeft:6}}>{it.cat}</span></div>
              <span style={{fontSize:13,fontWeight:700,color:C.blue}}>₹{selParty?getBestPurchRate(selParty.name,it.id,it.purch):it.purch}/pc</span>
            </div>)}
          </div>
        </div>}
      </div>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>{['#','Item','Qty','Rate ₹','Freight ₹/p','Amount',''].map((h,i)=><th key={i} style={{padding:'7px 12px',textAlign:i>1&&i<6?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}</tr></thead>
        <tbody>
          {!lines.length&&<tr><td colSpan={7} style={{padding:'20px',textAlign:'center',color:C.muted}}>No items</td></tr>}
          {lines.map((l,i)=><tr key={l.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
            <td style={{padding:'7px 12px',color:C.muted,fontSize:11}}>{i+1}</td>
            <td style={{padding:'7px 12px',fontWeight:600,color:C.txt}}>{l.supplierName||l.itemName}</td>
            <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.qty} onChange={e=>editLine(i,'qty',e.target.value)} type="number" style={{width:60,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
            <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.rate} onChange={e=>editLine(i,'rate',e.target.value)} type="number" style={{width:75,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
            <td style={{padding:'7px 12px',textAlign:'right'}}><input value={l.freightPerPeti||''} onChange={e=>editLine(i,'freightPerPeti',e.target.value)} type="number" placeholder="0" style={{width:70,padding:'3px 7px',borderRadius:5,border:`1px solid ${C.bdr}`,fontSize:12,textAlign:'right',outline:'none',fontFamily:'inherit'}}/></td>
            <td style={{padding:'7px 12px',textAlign:'right',fontWeight:700,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(l.qty*l.rate)}</td>
            <td style={{padding:'7px 12px',textAlign:'center'}}><button onClick={()=>setLines(p=>p.filter((_,j)=>j!==i))} style={{background:'none',border:'none',cursor:'pointer',color:C.muted,fontSize:15}}>×</button></td>
          </tr>)}
        </tbody>
        {lines.length>0&&<tfoot><tr style={{background:'#F6F8FA'}}><td colSpan={5} style={{padding:'9px 12px',fontWeight:700,textAlign:'right',color:C.txt}}>TOTAL</td><td style={{padding:'9px 12px',textAlign:'right',fontSize:16,fontWeight:800,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(total)}</td><td/></tr></tfoot>}
      </table>
    </div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'12px 16px',marginBottom:14}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:extraCharges.length?8:0}}>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5}}>Extra Charges <span style={{fontWeight:500,color:C.muted,textTransform:'none'}}>(loading, octroi, packing…)</span></label>
        <Btn onClick={addExtraCharge} variant="ghost" size="sm">+ Add Charge</Btn>
      </div>
      {extraCharges.map((c,i)=><div key={c.id} style={{display:'flex',gap:8,alignItems:'center',marginBottom:6}}>
        <input value={c.name} onChange={e=>editExtraCharge(i,'name',e.target.value)} list="extra-charge-names-edit" placeholder="Charge name — e.g. Loading"
          style={{flex:1,padding:'7px 10px',borderRadius:7,border:`1px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit'}}/>
        <input value={c.amount} onChange={e=>editExtraCharge(i,'amount',e.target.value)} type="number" placeholder="Amount"
          style={{width:110,padding:'7px 10px',borderRadius:7,border:`1px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',textAlign:'right'}}/>
        <button onClick={()=>removeExtraCharge(i)} style={{background:'none',border:'none',cursor:'pointer',color:C.muted,fontSize:16}}>×</button>
      </div>)}
      <datalist id="extra-charge-names-edit">{extraChargeNames.map(n=><option key={n} value={n}/>)}</datalist>
      {extraChargesTotal>0&&<div style={{fontSize:12,color:C.sub,textAlign:'right'}}>Extra charges total: <b style={{color:C.txt}}>{fmt(extraChargesTotal)}</b></div>}
    </div>
    {(freightAmount>0||extraChargesTotal>0)&&<div style={{fontSize:12,color:C.sub,marginBottom:14,textAlign:'right'}}>Goods {fmt(total)}{extraChargesTotal>0?' + Charges '+fmt(extraChargesTotal):''}{freightAmount>0?' + Freight '+fmt(freightAmount):''} = <b style={{color:C.txt}}>{fmt(total+extraChargesTotal+freightAmount)}</b></div>}
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <button onClick={()=>{if(window.confirm('Delete this purchase?')){
          if(purchase.receivedDate)(purchase.items||[]).forEach(l=>adjustStock(l.itemId,-l.qty));
          state.purchases=state.purchases.filter(p=>p.id!==purchase.id);save.purchases();
          state.freightCharges=state.freightCharges.filter(f=>f.purchaseId!==purchase.id);save.freightCharges();
          const linkedTxn=state.txns.find(t=>t.purchaseId===purchase.id);
          if(linkedTxn){state.txns=state.txns.filter(t=>t.id!==linkedTxn.id);save.txns();}
          state.supplierPayments=state.supplierPayments.filter(sp=>sp.purchaseId!==purchase.id);save.supplierPayments();
          onSaved?.();onClose();toast('Deleted');
        }}} style={{padding:'7px 14px',borderRadius:7,border:`1px solid ${C.red}`,background:C.redSoft,color:C.red,cursor:'pointer',fontSize:12,fontWeight:600,fontFamily:'inherit'}}>🗑 Delete</button>
      <div style={{display:'flex',gap:8}}>
        <Btn onClick={onClose} variant="secondary">Cancel</Btn>
        <Btn onClick={doSave} disabled={!lines.length} icon="💾">Save Changes</Btn>
      </div>
    </div>
  </Modal>;
}

// ── PURCHASE RETURN MODAL ──────────────────────────────────────────────────────
function PurchaseReturnModal({purchase,onClose,onSaved}){
  const normItem=(it,i)=>Array.isArray(it)
    ?{id:'RL'+i,itemId:'',itemName:it[0]||'',supplierName:it[0]||'',qty:Number(it[1])||0,rate:Number(it[2])||0,returnQty:Number(it[1])||0,returning:true}
    :{id:'RL'+i,itemId:it.itemId||'',itemName:it.supplierName||it.itemName||'',supplierName:it.supplierName||it.itemName||'',qty:Number(it.qty)||0,rate:Number(it.rate)||0,returnQty:Number(it.qty)||0,returning:true};
  const initLines=(purchase?.items||[]).map(normItem);
  const[returnLines,setReturnLines]=useState(initLines);
  const[date,setDate]=useState(today());
  const[reason,setReason]=useState('');
  const[party,setParty]=useState(state.parties.find(p=>p.id===purchase?.partyId)||null);
  const[pq,setPq]=useState('');
  const[selInv,setSelInv]=useState(null);
  const{toast,TC}=useToast();
  const supps=useMemo(()=>state.parties.filter(p=>p.type==='supplier'),[state.parties.length]);
  const psugg=useMemo(()=>pq.length<1?[]:supps.filter(p=>p.name.toLowerCase().includes(pq.toLowerCase())).slice(0,6),[supps,pq]);
  const partyHistory=useMemo(()=>{if(!party)return[];const hist=SUPPLIER_HISTORY[party.name]||{};return(hist.s||[]).map((s,i)=>({idx:i,date:s[0],total:s[1],items:s[2]||[]}));},[party]);
  const loadInvoice=inv=>{setSelInv(inv);setReturnLines(inv.items.map((x,i)=>({id:'RL'+i,itemId:'',itemName:Array.isArray(x)?x[0]:x,supplierName:Array.isArray(x)?x[0]:x,qty:Array.isArray(x)?x[1]:1,rate:Array.isArray(x)?x[2]:0,returnQty:Array.isArray(x)?x[1]:1,returning:true})));};
  const selected=returnLines.filter(l=>l.returning&&Number(l.returnQty)>0);
  const returnAmt=selected.reduce((s,l)=>s+(Number(l.returnQty)||0)*(Number(l.rate)||0),0);
  const doSave=()=>{
    if(!party)return toast('Select a supplier','error');
    if(!selected.length)return toast('Select at least one item','error');
    const ret={id:uid('PUR'),type:'return',date,bookedDate:date,receivedDate:date,partyId:party.id,partyName:party.name,refPurchaseId:purchase?.id||null,reason,
      items:selected.map(l=>({itemId:l.itemId,itemName:l.itemName,customerName:l.supplierName||l.itemName,supplierName:l.supplierName||l.itemName,qty:-(Number(l.returnQty)||0),rate:Number(l.rate)||0,amount:-(Number(l.returnQty)||0)*(Number(l.rate)||0)})),
      total:-returnAmt,createdAt:new Date().toISOString()};
    state.purchases=[ret,...state.purchases];save.purchases();
    selected.forEach(l=>adjustStock(l.itemId,-(Number(l.returnQty)||0)));
    toast('Return recorded — '+fmt(returnAmt)+' debited');
    onSaved?.();onClose();
  };
  return <Modal title="Record Purchase Return" onClose={onClose} width={560}>
    <TC/>
    {purchase&&<div style={{background:C.blueSoft,borderRadius:8,padding:'10px 14px',marginBottom:16,border:`1px solid ${C.blue}20`,fontSize:12}}>
      <span style={{color:C.sub}}>Returning against: </span><span style={{fontWeight:700,color:C.txt}}>{purchase.partyName}</span><span style={{color:C.sub,marginLeft:8}}>{fmtD(purchase.date)} · {fmt(purchase.total)}</span>
    </div>}
    <Grid2>
      <Inp label="Return Date" value={date} onChange={setDate} type="date" required/>
      <div>
        <label style={{fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,display:'block',marginBottom:3}}>Supplier</label>
        {party
          ?<div style={{display:'flex',alignItems:'center',gap:8,padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.blue}`,background:C.blueSoft}}>
            <span style={{fontWeight:700,flex:1,color:C.txt,fontSize:13}}>{party.name}</span>
            <Btn onClick={()=>{setParty(null);setPq('');setReturnLines([]);}} variant="ghost" size="sm">Change</Btn>
          </div>
          :<div style={{position:'relative'}}>
            <input value={pq} autoComplete="off" onChange={e=>setPq(e.target.value)} placeholder="Type supplier…"
              style={{width:'100%',padding:'7px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,color:C.txt,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}
              onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
            {psugg.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:C.surf,border:`1px solid ${C.bdr}`,borderRadius:'0 0 8px 8px',zIndex:100,boxShadow:'0 8px 24px rgba(0,0,0,0.12)',overflow:'hidden'}}>
              {psugg.map(p=><div key={p.id} onMouseDown={()=>{setParty(p);setPq('');}} style={{padding:'9px 14px',cursor:'pointer',fontSize:13,borderBottom:`1px solid ${C.bdr}`}} onMouseEnter={e=>e.currentTarget.style.background='#F9FAFB'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}><span style={{fontWeight:600,color:C.txt}}>{p.name}</span></div>)}
            </div>}
          </div>}
      </div>
      <div style={{gridColumn:'1/-1'}}><Inp label="Reason (optional)" value={reason} onChange={setReason} placeholder="e.g. Damaged, Wrong item"/></div>
    </Grid2>
    {party&&partyHistory.length>0&&!purchase&&<div style={{marginBottom:14}}>
      <div style={{fontSize:12,fontWeight:700,color:C.txt,marginBottom:8}}>Select Purchase to Return Against</div>
      <div style={{maxHeight:160,overflowY:'auto',borderRadius:8,border:`1px solid ${C.bdr}`}}>
        {partyHistory.map((inv,i)=><div key={i} onClick={()=>loadInvoice(inv)}
          style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'9px 14px',cursor:'pointer',borderBottom:i<partyHistory.length-1?`1px solid ${C.bdr}`:'none',background:selInv?.idx===inv.idx?C.blueSoft:'transparent'}}
          onMouseEnter={e=>e.currentTarget.style.background=C.blueSoft} onMouseLeave={e=>e.currentTarget.style.background=selInv?.idx===inv.idx?C.blueSoft:'transparent'}>
          <div><span style={{fontWeight:600,color:C.txt,fontSize:13}}>{fmtD(inv.date)}</span><span style={{fontSize:11,color:C.sub,marginLeft:10}}>{inv.items.slice(0,2).map(x=>Array.isArray(x)?x[0]:x).join(', ')}{inv.items.length>2?` +${inv.items.length-2} more`:''}</span></div>
          <span style={{fontWeight:700,color:C.blue,fontVariantNumeric:'tabular-nums'}}>{fmt(inv.total)}</span>
        </div>)}
      </div>
      {selInv&&<div style={{fontSize:11,color:C.green,marginTop:6,fontWeight:600}}>✓ Items loaded from {fmtD(selInv.date)} — adjust quantities below</div>}
    </div>}
    <div style={{margin:'14px 0 4px',fontSize:12,fontWeight:700,color:C.txt}}>Items to Return</div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',marginBottom:16}}>
      {!returnLines.length&&<div style={{padding:20,textAlign:'center',color:C.muted,fontSize:13}}>{party?'Select a purchase above to load items':'Select a supplier first'}</div>}
      {returnLines.map((l,i)=><div key={l.id} style={{display:'flex',alignItems:'center',gap:10,padding:'10px 14px',borderBottom:i<returnLines.length-1?`1px solid ${C.bdr}`:'none',opacity:l.returning?1:.6}}>
        <input type="checkbox" checked={l.returning} onChange={e=>setReturnLines(p=>p.map((x,j)=>j!==i?x:{...x,returning:e.target.checked}))} style={{cursor:'pointer',width:15,height:15,flexShrink:0}}/>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontWeight:600,color:C.txt,fontSize:13,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{l.supplierName||l.itemName}</div>
          <div style={{fontSize:11,color:C.sub}}>Original: {l.qty} pcs @ {fmt(l.rate)}</div>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:6,flexShrink:0}}>
          <span style={{fontSize:11,color:C.sub}}>Return qty:</span>
          <input value={l.returnQty} onChange={e=>setReturnLines(p=>p.map((x,j)=>j!==i?x:{...x,returnQty:Number(e.target.value)||0}))} type="number" disabled={!l.returning} min={0} max={l.qty}
            style={{width:65,padding:'4px 8px',borderRadius:6,border:`1.5px solid ${l.returning?C.bdr:'transparent'}`,fontSize:12,textAlign:'center',outline:'none',fontFamily:'inherit'}}/>
          <span style={{fontSize:13,fontWeight:700,color:C.red,minWidth:70,textAlign:'right',fontVariantNumeric:'tabular-nums'}}>{l.returning&&Number(l.returnQty)>0?fmt(Number(l.returnQty)*Number(l.rate)):''}</span>
        </div>
      </div>)}
    </div>
    {returnAmt>0&&<div style={{background:C.redSoft,borderRadius:8,border:`1px solid ${C.red}30`,padding:'12px 16px',marginBottom:16,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <div><div style={{fontSize:12,color:C.sub}}>Return Amount</div><div style={{fontSize:22,fontWeight:800,color:C.red,fontVariantNumeric:'tabular-nums'}}>{fmt(returnAmt)}</div></div>
      <div style={{fontSize:12,color:C.sub,textAlign:'right'}}><div>{selected.length} items · {selected.reduce((s,l)=>s+Number(l.returnQty),0)} pcs</div><div style={{marginTop:2}}>Will reduce payable by {fmt(returnAmt)}</div></div>
    </div>}
    <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
      <Btn onClick={onClose} variant="secondary">Cancel</Btn>
      <Btn onClick={doSave} disabled={!selected.length||returnAmt===0||!party} variant="danger">Record Return · −{fmt(returnAmt)}</Btn>
    </div>
  </Modal>;
}

function PurchLog() {
  const[search,setSearch]=useState('');
  const[dateFrom,setDateFrom]=useState('');
  const[dateTo,setDateTo]=useState('');
  const[page,setPage]=useState(0);
  const[_fvPL,_setFvPL]=useState(0); const forceUpdate=()=>_setFvPL(v=>v+1);
  const{toast,TC}=useToast();
  const[ledgerParty,setLedgerParty]=useState(null);
  const[viewPurchase,setViewPurchase]=useState(null);
  const[editPurchase,setEditPurchase]=useState(null);
  const[retPurchase,setRetPurchase]=useState(false);
  const PAGE=50;

  const markReceived=id=>{
    const p=state.purchases.find(x=>x.id===id);
    if(!p||p.receivedDate)return;
    const rd=today();
    state.purchases=state.purchases.map(x=>x.id===id?{...x,receivedDate:rd}:x);
    save.purchases();
    (p.items||[]).forEach(l=>adjustStock(l.itemId,l.qty));
    forceUpdate();
    toast('Marked received — stock updated ✓');
  };

  // Historical (raw Excel Purchase sheet, via SUPPLIER_HISTORY) + new ERP purchases
  const allPurchases=useMemo(()=>{
    const rows=[];
    for(const[vendor,hist] of Object.entries(SUPPLIER_HISTORY)){
      for(const s of (hist.s||[])){
        rows.push({d:s[0],receivedDate:s[0],party:vendor,total:s[1],items:s[2]||[]});
      }
    }
    for(const p of state.purchases){
      rows.push({d:p.bookedDate||p.date,receivedDate:p.receivedDate||null,party:p.partyName,total:p.total,items:(p.items||[]).map(x=>([x.supplierName||x.customerName||x.itemName,x.qty,x.rate])),isNew:true,isReturn:p.total<0,id:p.id});
    }
    return rows.sort((a,b)=>b.d.localeCompare(a.d));
  },[state.purchases.length,_fvPL]);

  const filtered=useMemo(()=>allPurchases.filter(r=>
    (!search||r.party.toLowerCase().includes(search.toLowerCase())||r.items.some(x=>(x[0]||'').toLowerCase().includes(search.toLowerCase())))&&
    (!dateFrom||r.d>=dateFrom)&&(!dateTo||r.d<=dateTo)
  ),[allPurchases,search,dateFrom,dateTo]);

  const pageData=filtered.slice(page*PAGE,(page+1)*PAGE);
  const totalAmt=filtered.reduce((s,r)=>s+r.total,0);
  const totalPages=Math.ceil(filtered.length/PAGE);

  const viewRow=r=>{
    if(r.isNew){
      const found=state.purchases.find(x=>x.id===r.id);
      if(found)setViewPurchase(found);
    }else{
      const p=state.parties.find(pp=>pp.name===r.party);
      setViewPurchase({id:null,date:r.d,bookedDate:r.d,receivedDate:r.receivedDate||null,partyId:p?.id,partyName:r.party,total:r.total,
        items:(r.items||[]).map(x=>({itemId:null,itemName:x[0],customerName:x[0],supplierName:x[0],qty:x[1],rate:x[2],amount:x[1]*x[2]})),freightAmount:0});
    }
  };

  return <div>
    <TC/>
    {ledgerParty&&<PartyLedger party={ledgerParty} onClose={()=>setLedgerParty(null)}/>}
    {viewPurchase&&<PurchaseViewModal purchase={viewPurchase} onClose={()=>setViewPurchase(null)}/>}
    {editPurchase&&<EditPurchaseModal purchase={editPurchase} onClose={()=>setEditPurchase(null)} onSaved={forceUpdate}/>}
    {retPurchase!==false&&<PurchaseReturnModal purchase={retPurchase||null} onClose={()=>setRetPurchase(false)} onSaved={forceUpdate}/>}
    <SectionHead title="Purchases" sub={`${filtered.filter(r=>!r.isReturn).length} entries · ${fmtL(filtered.filter(r=>!r.isReturn).reduce((s,r)=>s+r.total,0))}${filtered.filter(r=>r.isReturn).length>0?' · '+filtered.filter(r=>r.isReturn).length+' returns −'+fmtL(Math.abs(filtered.filter(r=>r.isReturn).reduce((s,r)=>s+r.total,0))):''}`}
      action={<Btn onClick={()=>setRetPurchase(null)} variant='danger' size='sm'>+ New Return</Btn>}/>
    <PurchaseEntry onSaved={forceUpdate}/>
    <div style={{display:'flex',gap:10,marginBottom:14,flexWrap:'wrap'}}>
      <input value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}} placeholder="Search supplier or item…"
        style={{flex:1,minWidth:200,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}
        onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
      <input type="date" value={dateFrom} onChange={e=>{setDateFrom(e.target.value);setPage(0);}}
        style={{padding:'8px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
      <input type="date" value={dateTo} onChange={e=>{setDateTo(e.target.value);setPage(0);}}
        style={{padding:'8px 10px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
      {(search||dateFrom||dateTo)&&<Btn onClick={()=>{setSearch('');setDateFrom('');setDateTo('');setPage(0);}} variant="secondary" size="sm">Clear</Btn>}
    </div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>
          {['Booked','Received','Supplier','Item / Model','Amount',''].map((h,i)=><th key={i} style={{padding:'9px 14px',textAlign:i===4?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}
        </tr></thead>
        <tbody>
          {pageData.length===0&&<tr><td colSpan={6} style={{padding:'40px',textAlign:'center',color:C.muted}}>No purchases recorded yet</td></tr>}
          {pageData.map((r,i)=><tr key={i} onClick={()=>viewRow(r)} style={{borderBottom:`1px solid ${C.bdr}`,cursor:'pointer',background:r.isReturn?C.redSoft:r.isNew?C.blueSoft:'transparent'}}
            onMouseEnter={e=>e.currentTarget.style.background=r.isReturn?C.redSoft:r.isNew?C.blueSoft:'#F9FAFB'}
            onMouseLeave={e=>e.currentTarget.style.background=r.isReturn?C.redSoft:r.isNew?C.blueSoft:'transparent'}>
            <td style={{padding:'9px 14px',whiteSpace:'nowrap',color:C.sub}}>{fmtD(r.d)}{r.isNew&&<span style={{fontSize:10,color:C.blue,fontWeight:700,marginLeft:5}}>NEW</span>}</td>
            <td style={{padding:'9px 14px',whiteSpace:'nowrap'}}>{r.receivedDate?<span style={{color:C.sub}}>{fmtD(r.receivedDate)}</span>:<Badge v="Pending" color="amber"/>}</td>
            <td style={{padding:'9px 14px',fontWeight:600,color:C.txt,maxWidth:200,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
              <span onClick={e=>{e.stopPropagation();const p=state.parties.find(pp=>pp.name===r.party);if(p)setLedgerParty(p);}}
                style={{cursor:'pointer',color:C.blue,textDecoration:'underline',textDecorationStyle:'dotted'}}>{r.party}</span>
            </td>
            <td style={{padding:'9px 14px',fontSize:12,color:C.sub}}>
              {r.items.length>0?r.items.slice(0,2).map((x,j)=><span key={j} style={{display:'block'}}>{x[0]} ×{x[1]} @₹{x[2]}</span>):'—'}
              {r.items.length>2&&<span style={{color:C.muted,fontSize:11}}>+{r.items.length-2} more</span>}
            </td>
            <td style={{padding:'9px 14px',textAlign:'right',fontWeight:700,color:r.isReturn?C.red:C.blue,fontVariantNumeric:'tabular-nums'}}>{r.isReturn&&'−'}{fmt(Math.abs(r.total))}</td>
            <td style={{padding:'9px 14px',textAlign:'center'}}>
              <div style={{display:'flex',gap:5,justifyContent:'center'}}>
                {r.isNew&&!r.receivedDate&&!r.isReturn&&<Btn onClick={e=>{e.stopPropagation();markReceived(r.id);}} size="sm" variant="success">Mark Received</Btn>}
                <IconBtn title="View purchase" color={C.blue} bg={C.blueSoft} onClick={e=>{e.stopPropagation();viewRow(r);}}>👁</IconBtn>
                {r.isNew&&<IconBtn title="Edit" color={C.sub} bg={C.surf}
                  onClick={e=>{e.stopPropagation();const found=state.purchases.find(x=>x.id===r.id);if(found)setEditPurchase(found);}}>✎</IconBtn>}
                {!r.isReturn&&<IconBtn title="Record return" color={C.red} bg={C.redSoft}
                  onClick={e=>{e.stopPropagation();
                    if(r.isNew){const found=state.purchases.find(x=>x.id===r.id);setRetPurchase(found||null);}
                    else setRetPurchase({partyName:r.party,partyId:state.parties.find(p=>p.name===r.party)?.id||'',date:r.d,total:r.total,items:r.items||[]});
                  }}>↩</IconBtn>}
                {r.isNew&&<IconBtn title="Delete" color={C.red} bg={C.redSoft}
                  onClick={e=>{e.stopPropagation();
                    if(!window.confirm('Delete this purchase?'))return;
                    const found=state.purchases.find(x=>x.id===r.id);
                    if(found){
                      if(found.receivedDate)(found.items||[]).forEach(l=>adjustStock(l.itemId,-l.qty));
                      state.purchases=state.purchases.filter(x=>x.id!==found.id);save.purchases();
                      state.freightCharges=state.freightCharges.filter(f=>f.purchaseId!==found.id);save.freightCharges();
                      const linkedTxn=state.txns.find(t=>t.purchaseId===found.id);
                      if(linkedTxn){state.txns=state.txns.filter(t=>t.id!==linkedTxn.id);save.txns();}
                      state.supplierPayments=state.supplierPayments.filter(sp=>sp.purchaseId!==found.id);save.supplierPayments();
                      forceUpdate();toast('Deleted');
                    }
                  }}>🗑</IconBtn>}
              </div>
            </td>
          </tr>)}
        </tbody>
        <tfoot><tr style={{background:'#F6F8FA'}}>
          <td colSpan={4} style={{padding:'9px 14px',fontWeight:700,textAlign:'right',color:C.txt}}>Total</td>
          <td style={{padding:'9px 14px',textAlign:'right',fontWeight:800,color:C.blue}}>{fmt(totalAmt)}</td>
          <td/>
        </tr></tfoot>
      </table>
    </div>
    {totalPages>1&&<div style={{display:'flex',alignItems:'center',justifyContent:'center',gap:8,marginTop:14}}>
      <Btn onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} variant="secondary" size="sm">← Prev</Btn>
      <span style={{fontSize:12,color:C.sub}}>Page {page+1} of {totalPages}</span>
      <Btn onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={page>=totalPages-1} variant="secondary" size="sm">Next →</Btn>
    </div>}
  </div>;
}

// ── BACKUP & RESTORE ─────────────────────────────────────────────────────────
// Backs up/restores every key `save` knows how to persist, rather than a
// hand-picked list — the old hand-picked list predated most of the app's
// data model (items, purchases, supplierPayments, freightCharges, orders,
// customCategories, categoryTags, cityTransport, dispatchPhone,
// itemCustomerNames, cityOrder, repairExpenses were all added later) and
// silently left every one of them out of every backup, so restoring one
// would wipe out inventory, purchase history, custom categories and more.
// It also referenced `DB` directly without importing it, so Export crashed
// immediately with a ReferenceError before writing anything. Driving both
// backup and restore off Object.keys(save) means a future new piece of
// persisted state is automatically included without this file needing a
// matching update.
function BackupRestore() {
  const [restoring, setRestoring] = useState(false);
  const [status, setStatus]       = useState('');
  const [preview, setPreview]     = useState(null);   // parsed backup before confirming
  const { toast, TC } = useToast();
  const STATE_KEYS = Object.keys(save);

  // ── BACKUP ────────────────────────────────────────────────────────────────
  const doBackup = async () => {
    try {
      const payload = {};
      for (const k of STATE_KEYS) payload['data:'+k] = state[k];

      const backup = {
        _app:       'Upkaran ERP',
        _ver:       APP_VERSION,
        _date:      new Date().toISOString(),
        _parties:   state.parties.length,
        _items:     state.items.length,
        _sales:     state.sales.length,
        _purchases: state.purchases.length,
        _pmts:      state.payments.length,
        _fus:       state.followups.length,
        _phones:    Object.keys(state.phoneEdits).length,
        data:       payload,
      };

      const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = 'upkaran-backup-' + today() + '.json';
      document.body.appendChild(a); a.click();
      document.body.removeChild(a); URL.revokeObjectURL(url);
      toast('Backup downloaded ✓');
    } catch (e) {
      toast('Backup failed: ' + e.message, 'error');
    }
  };

  // ── RESTORE ───────────────────────────────────────────────────────────────
  const onFileSelect = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result);
        if (parsed._app !== 'Upkaran ERP') {
          toast('Not a valid Upkaran backup file', 'error');
          return;
        }
        setPreview(parsed);
      } catch {
        toast('Invalid JSON file', 'error');
      }
    };
    reader.readAsText(file);
    e.target.value = '';   // reset so same file can be re-selected
  };

  const doRestore = async () => {
    if (!preview) return;
    setRestoring(true); setStatus('Restoring…');
    try {
      const d = preview.data;
      for (const k of STATE_KEYS) {
        const key = 'data:'+k;
        if (d[key] === undefined) continue;
        state[k] = k === 'templates' ? {...DEFAULT_TEMPLATES, ...d[key]} : d[key];
        await save[k]();
      }
      if (d['data:parties'] === undefined) {
        // Old backups predate party persistence and have no parties key at
        // all — keep whatever's currently loaded and at least carry over
        // phone edits from the backup, rather than losing parties entirely.
        state.parties = state.parties.map(p => state.phoneEdits[p.id] !== undefined ? {...p, phone: state.phoneEdits[p.id]} : p);
        await save.parties();
      }
      setStatus('Restore complete ✓');
      setPreview(null);
      toast('Restored successfully — reloading…');
      setTimeout(() => window.location.reload(), 1500);
    } catch (e) {
      toast('Restore failed: ' + e.message, 'error');
      setRestoring(false); setStatus('');
    }
  };

  const fmtDate = iso => iso ? new Date(iso).toLocaleString('en-IN', {day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}) : '—';

  return (
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <TC/>
      <div style={{padding:'14px 20px',borderBottom:`1px solid ${C.bdr}`,background:'#F6F8FA'}}>
        <h3 style={{margin:0,fontSize:15,fontWeight:700,color:C.txt}}>Backup & Restore</h3>
        <p style={{margin:'4px 0 0',fontSize:12,color:C.sub}}>Export everything — parties, inventory, sales, purchases, payments, follow-ups and settings — to a JSON file. Restore from any previous backup.</p>
      </div>

      {/* Backup section */}
      <div style={{padding:'18px 20px',borderBottom:`1px solid ${C.bdr}`}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
          <div>
            <div style={{fontSize:13,fontWeight:600,color:C.txt,marginBottom:2}}>Download Backup</div>
            <div style={{fontSize:12,color:C.sub}}>
              {state.parties.length} parties · {state.items.length} items · {state.sales.length} sales · {state.purchases.length} purchases · {state.payments.length} payments · {state.followups.length} follow-ups
            </div>
          </div>
          <Btn onClick={doBackup} icon="⬇">Export JSON</Btn>
        </div>
      </div>

      {/* Restore section */}
      <div style={{padding:'18px 20px'}}>
        <div style={{fontSize:13,fontWeight:600,color:C.txt,marginBottom:8}}>Restore from Backup</div>

        {!preview
          ? <div>
              <div style={{fontSize:12,color:C.sub,marginBottom:12}}>
                Select a <code style={{background:'#F3F4F6',padding:'1px 5px',borderRadius:3,fontSize:11}}>upkaran-backup-*.json</code> file to restore. This will overwrite current data.
              </div>
              <label style={{display:'inline-flex',alignItems:'center',gap:8,padding:'8px 18px',borderRadius:7,
                border:`1.5px solid ${C.bdr}`,background:C.surf,cursor:'pointer',fontSize:13,fontWeight:600,color:C.txt}}>
                📂 Choose backup file
                <input type="file" accept=".json" onChange={onFileSelect} style={{display:'none'}}/>
              </label>
            </div>
          : <div style={{background:'#FEF9C3',borderRadius:8,border:`1.5px solid #FDE047`,padding:16}}>
              <div style={{fontWeight:700,color:'#713F12',fontSize:13,marginBottom:10}}>⚠ Review before restoring</div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6,fontSize:12,marginBottom:14}}>
                {[
                  ['Backup date',    fmtDate(preview._date)],
                  ['App version',    preview._ver],
                  ['Parties',        (preview._parties??'—')+' records'],
                  ['Items',          (preview._items??'—')+' products'],
                  ['Sales',          preview._sales+' invoices'],
                  ['Purchases',      (preview._purchases??'—')+' records'],
                  ['Payments',       preview._pmts+' records'],
                  ['Follow-ups',     preview._fus+' entries'],
                  ['Phone edits',    preview._phones+' numbers'],
                ].map(([l,v])=><div key={l} style={{display:'flex',gap:6}}>
                  <span style={{color:'#92400E',fontWeight:500}}>{l}:</span>
                  <span style={{color:'#78350F',fontWeight:700}}>{v}</span>
                </div>)}
              </div>
              <div style={{fontSize:11,color:'#92400E',marginBottom:12}}>
                This will replace all current parties, items, sales, purchases, payments, follow-ups and settings.
              </div>
              <div style={{display:'flex',gap:8}}>
                <Btn onClick={doRestore} disabled={restoring} variant="danger">
                  {restoring ? status : 'Confirm Restore'}
                </Btn>
                <Btn onClick={()=>setPreview(null)} variant="secondary" disabled={restoring}>Cancel</Btn>
              </div>
            </div>
        }
      </div>
    </div>
  );
}

// Cloud sync status + the one-time "push this device's data up" action —
// used exactly once, from whichever device already holds the real
// accumulated business data, to seed the cloud table the first time this
// feature is turned on. Every other device's initSync() then pulls that
// real data instead of starting from seed/empty. Renders nothing (not even
// the card) when this build has no Supabase project configured.
function CloudSyncPanel(){
  const{toast,TC}=useToast();
  const[email,setEmail]=useState(null);
  const[busy,setBusy]=useState(false);
  useEffect(()=>{ DB.getSession().then(s=>setEmail(s?.user?.email||null)); },[]);
  if(!supabase)return null;
  const upload=async()=>{
    if(!window.confirm('Upload this device\'s current data to the cloud? Only do this from the device that already has your real, up-to-date data — it will overwrite the cloud copy.'))return;
    setBusy(true);
    try{ const n=await DB.pushAllLocal(); toast(`Uploaded ${n} data keys to the cloud ✓`); }
    catch(e){ toast('Upload failed: '+e.message,'error'); }
    finally{ setBusy(false); }
  };
  const signOut=async()=>{ await DB.signOut(); window.location.reload(); };
  return <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:20,marginBottom:20}}>
    <TC/>
    <h3 style={{margin:'0 0 4px',fontSize:14,fontWeight:700,color:C.txt}}>Cloud Sync</h3>
    <p style={{margin:'0 0 12px',fontSize:12,color:C.sub}}>Signed in as <strong style={{color:C.txt}}>{email||'…'}</strong> — this device's data syncs automatically across every signed-in device.</p>
    <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
      <Btn onClick={upload} disabled={busy} size="sm">{busy?'Uploading…':'Upload this device\'s data to cloud'}</Btn>
      <button onClick={signOut} style={{fontSize:12,color:C.sub,background:'none',border:`1px solid ${C.bdr}`,cursor:'pointer',fontFamily:'inherit',padding:'6px 12px',borderRadius:6}}>Sign out</button>
    </div>
  </div>;
}

// ── SETTINGS ──────────────────────────────────────────────────────────────────
function Settings(){
  const[tpls,setTpls]=useState({...state.templates});
  const[active,setActive]=useState('invoice');
  const[dispatchPhone,setDispatchPhone]=useState(state.dispatchPhone||'');
  const{toast,TC}=useToast();
  const VARS={invoice:['{date}','{partyName}','{transport}','{freight}','{items}','{total}','{credit}'],payment:['{partyName}','{amount}','{invoiceDetails}'],followup:['{partyName}','{lastItem}','{lastDate}'],delivery:['{partyName}','{items}','{lr}'],dispatch:['{date}','{partyName}','{city}','{items}','{totalQty}','{transport}']};
  const LABELS={invoice:'📄 Invoice',payment:'💰 Payment reminder',followup:'🔁 Follow-up',delivery:'🚚 Delivery confirmation',dispatch:'📦 Dispatch order'};
  const doSave=async()=>{state.templates={...state.templates,...tpls};await save.templates();toast('Templates saved ✓');};
  const reset=()=>setTpls({...DEFAULT_TEMPLATES});
  return <div style={{maxWidth:700}}>
    <TC/>
    <SectionHead title="Settings"/>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',marginBottom:20}}>
      <div style={{padding:'14px 20px',borderBottom:`1px solid ${C.bdr}`}}>
        <h3 style={{margin:0,fontSize:15,fontWeight:700,color:C.txt}}>WhatsApp Message Templates</h3>
        <p style={{margin:'4px 0 0',fontSize:12,color:C.sub}}>Click a variable chip to insert it at cursor.</p>
      </div>
      <div style={{display:'flex',borderBottom:`1px solid ${C.bdr}`,background:'#F6F8FA'}}>
        {Object.entries(LABELS).map(([key,label])=>(
          <button key={key} onClick={()=>setActive(key)} style={{padding:'10px 18px',border:'none',background:'transparent',cursor:'pointer',fontFamily:'inherit',fontSize:13,fontWeight:active===key?700:400,color:active===key?C.blue:C.sub,borderBottom:active===key?`2px solid ${C.blue}`:'2px solid transparent',marginBottom:-1}}>{label}</button>
        ))}
      </div>
      <div style={{padding:20}}>
        <div style={{display:'flex',gap:6,flexWrap:'wrap',marginBottom:12}}>
          <span style={{fontSize:11,color:C.sub,fontWeight:600,alignSelf:'center'}}>Insert:</span>
          {VARS[active].map(v=>(
            <button key={v} onClick={()=>setTpls(t=>({...t,[active]:t[active]+v}))} style={{padding:'3px 10px',borderRadius:20,fontSize:11,fontWeight:600,background:C.blueSoft,color:C.blue,border:`1px solid ${C.blueMid}`,cursor:'pointer',fontFamily:'monospace'}}>{v}</button>
          ))}
        </div>
        <textarea value={tpls[active]||''} onChange={e=>setTpls(t=>({...t,[active]:e.target.value}))} rows={10}
          style={{width:'100%',padding:'12px 14px',borderRadius:8,border:`1.5px solid ${C.bdr}`,fontSize:13,fontFamily:'monospace',lineHeight:1.7,outline:'none',resize:'vertical',boxSizing:'border-box',color:C.txt}}
          onFocus={e=>e.target.style.borderColor=C.blue} onBlur={e=>e.target.style.borderColor=C.bdr}/>
        <div style={{marginTop:10,padding:'11px 14px',background:C.greenSoft,borderRadius:8,border:`1px solid ${C.green}30`,fontSize:12,fontFamily:'monospace',lineHeight:1.8,whiteSpace:'pre-wrap',color:C.txt}}>
          {applyTemplate(active,{date:'11 Jul 2026',partyName:'Ravi Traders, Sakti',transport:'*Transport:* Mahamaya | LR: 12345',freight:'',items:'• Infrared x36 pcs @ Rs1,900 = *Rs68,400*',total:'Rs68,400',credit:'Credit 30d | Due by 10 Aug 2026',amount:'Rs74,400',lastItem:'Infrared',lastDate:'2 Jul 2026',invoiceDetails:'• 2 Jul 2026 — Rs68,400 (9d from invoice)\n• 6 Jun 2026 — Rs6,000 (35d from invoice)',lr:'LR: 12345',city:'Sakti',totalQty:'36'})}
        </div>
        <div style={{fontSize:11,color:C.muted,marginTop:4}}>↑ Live preview with sample data</div>
      </div>
      <div style={{padding:'12px 20px',borderTop:`1px solid ${C.bdr}`,display:'flex',gap:10,justifyContent:'space-between',alignItems:'center'}}>
        <button onClick={reset} style={{fontSize:12,color:C.sub,background:'none',border:`1px solid ${C.bdr}`,cursor:'pointer',fontFamily:'inherit',padding:'6px 12px',borderRadius:6}}>Reset to defaults</button>
        <Btn onClick={doSave}>Save Templates</Btn>
      </div>
    </div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:20,marginBottom:20}}>
      <h3 style={{margin:'0 0 4px',fontSize:14,fontWeight:700,color:C.txt}}>Dispatch WhatsApp Number</h3>
      <p style={{margin:'0 0 12px',fontSize:12,color:C.sub}}>Used by Sales Order's "Send to Dispatch" step — your warehouse/dispatch team's number.</p>
      <div style={{display:'flex',gap:8}}>
        <input value={dispatchPhone} onChange={e=>setDispatchPhone(e.target.value)} placeholder="10-digit WhatsApp number" autoComplete="off"
          style={{flex:1,padding:'8px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit'}}/>
        <Btn onClick={()=>{if(!dispatchPhone||dispatchPhone.length<10)return toast('Enter 10-digit number','error');state.dispatchPhone=dispatchPhone;save.dispatchPhone();toast('Saved ✓');}} size="sm">Save</Btn>
      </div>
    </div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:20}}>
      <h3 style={{margin:'0 0 12px',fontSize:14,fontWeight:700,color:C.txt}}>About</h3>
      <div style={{fontSize:13,color:C.sub,lineHeight:1.9}}>
        <div>Business: <strong style={{color:C.txt}}>Upkaran, Raipur</strong></div>
        <div>Owner: <strong style={{color:C.txt}}>Pratik Agrawal</strong></div>
        <div>Data cut-off: <strong style={{color:C.txt}}>Jul 11, 2026</strong></div>
        <div>App version: <strong style={{color:C.txt}}>{APP_VERSION}</strong></div>
        <div>Parties: <strong style={{color:C.txt}}>{state.parties.filter(p=>p.type==='customer').length} customers</strong></div>
        <div>Items: <strong style={{color:C.txt}}>{state.items.length} products</strong></div>
      </div>
    </div>

    <div style={{height:16}}/>
    <CloudSyncPanel/>
    <BackupRestore/>
  </div>;
}

// ── REPORTS (Expenses / P&L / Balance Sheet) ───────────────────────────────────
function ExpensesReport(){
  const[_fv,_setFv]=useState(0); const forceUpdate=()=>_setFv(v=>v+1);
  const[ledgerCat,setLedgerCat]=useState(null);
  const rows=useMemo(()=>getExpenseReport(),[_fv,state.txns.length,state.customCategories.length]);
  const bizTotal=rows.filter(r=>r.biz==='business').reduce((s,r)=>s+r.total,0);
  const persTotal=rows.filter(r=>r.biz!=='business').reduce((s,r)=>s+r.total,0);
  const toggle=id=>{const cur=rows.find(r=>r.id===id).biz;setCategoryBiz(id,cur==='business'?'personal':'business');forceUpdate();};
  return <div>
    {ledgerCat&&<CategoryLedger catId={ledgerCat} onClose={()=>setLedgerCat(null)}/>}
    <SectionHead title="Expenses" sub={`${fmtL(bizTotal+persTotal)} total · ${fmtL(bizTotal)} business · ${fmtL(persTotal)} personal`}/>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden',overflowX:'auto'}}>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>
          {['Category','Historical (pre-app)','New (this app)','Total','Tag'].map((h,i)=><th key={i} style={{padding:'9px 12px',textAlign:i>0&&i<4?'right':i===4?'center':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`,whiteSpace:'nowrap'}}>{h}</th>)}
        </tr></thead>
        <tbody>
          {rows.map(r=><tr key={r.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
            <td style={{padding:'9px 12px',fontWeight:600,color:C.blue,cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted'}} onClick={()=>setLedgerCat(r.id)}>{r.label}</td>
            <td style={{padding:'9px 12px',textAlign:'right',color:C.sub}}>{fmt(r.historical)}</td>
            <td style={{padding:'9px 12px',textAlign:'right',color:C.sub}}>{fmt(r.newAmt)}</td>
            <td style={{padding:'9px 12px',textAlign:'right',fontWeight:700,color:C.txt}}>{fmt(r.total)}</td>
            <td style={{padding:'9px 12px',textAlign:'center'}}>
              <button onClick={()=>toggle(r.id)} style={{padding:'3px 10px',borderRadius:20,border:'none',cursor:'pointer',fontSize:11,fontWeight:700,fontFamily:'inherit',
                background:r.biz==='business'?C.blueSoft:C.amberSoft,color:r.biz==='business'?C.blue:C.amber}}>
                {r.biz==='business'?'Business':'Personal'}
              </button>
            </td>
          </tr>)}
        </tbody>
        <tfoot><tr style={{background:'#F6F8FA'}}>
          <td style={{padding:'9px 12px',fontWeight:700,color:C.txt}}>Total</td>
          <td style={{padding:'9px 12px',textAlign:'right',fontWeight:700}}>{fmt(rows.reduce((s,r)=>s+r.historical,0))}</td>
          <td style={{padding:'9px 12px',textAlign:'right',fontWeight:700}}>{fmt(rows.reduce((s,r)=>s+r.newAmt,0))}</td>
          <td style={{padding:'9px 12px',textAlign:'right',fontWeight:800}}>{fmt(bizTotal+persTotal)}</td>
          <td/>
        </tr></tfoot>
      </table>
    </div>
    <div style={{fontSize:11,color:C.muted,marginTop:10}}>Click a category name to see its full ledger (in and out). Click a tag to reclassify — this feeds the P&L (only Business expenses reduce Net Profit).</div>
  </div>;
}

function PnLReport(){
  const pnl=useMemo(()=>getPnL(),[state.txns.length,state.sales.length,state.customCategories.length]);
  const Row=({label,val,bold,indent,color})=><div style={{display:'flex',justifyContent:'space-between',padding:'8px 0',borderBottom:`1px solid ${C.bdr}`,fontWeight:bold?800:500,fontSize:bold?14:13,paddingLeft:indent?16:0}}>
    <span style={{color:C.txt}}>{label}</span><span style={{color:color||C.txt,fontVariantNumeric:'tabular-nums'}}>{val<0?'−':''}{fmt(Math.abs(val))}</span>
  </div>;
  return <div>
    <SectionHead title="Profit & Loss" sub="All-time — historical (pre-app) + recorded in this app"/>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 22px',maxWidth:560}}>
      <Row label="Revenue (Sales)" val={pnl.revenue} color={C.blue}/>
      <Row label="Cost of Goods Sold" val={-pnl.cogs} indent color={C.sub}/>
      <Row label="Gross Profit" val={pnl.grossProfit} bold color={pnl.grossProfit>=0?C.green:C.red}/>
      <Row label="Operating Expenses (Business)" val={-pnl.businessExpenses} indent color={C.sub}/>
      {pnl.repairExpenseTotal>0&&<Row label="Repair Expenditure (Warranty)" val={-pnl.repairExpenseTotal} indent color={C.purple}/>}
      <Row label="Net Profit" val={pnl.netProfit} bold color={pnl.netProfit>=0?C.green:C.red}/>
      <div style={{marginTop:14,paddingTop:14,borderTop:`2px dashed ${C.bdr}`}}>
        <Row label="Non-Business / Personal Spend (excluded above)" val={pnl.nonBusinessExpenses} color={C.amber}/>
      </div>
    </div>
    <div style={{fontSize:11,color:C.muted,marginTop:10,maxWidth:560}}>Historical COGS is each item's actual purchase rate (excluding transport, which has its own Expenses line) at the time it sold. New sales recorded in this app use the item's current purchase rate as a stand-in until real per-sale cost is tracked.</div>
  </div>;
}

// Category-wise stock breakdown, opened from the Balance Sheet's "Inventory"
// line — mirrors the grouping CategoryAnalytics uses (item.cat||'Other'),
// but limited to count/pcs/value since this is an assets-snapshot context,
// not a sales/profit report. Clicking a row reuses the existing CategoryDetail
// modal for item-level drill-down rather than duplicating that table.
function StockBreakdownModal({onClose}){
  const[catDetail,setCatDetail]=useState(null);
  // FIFO-costed value (see getFifoStockMap in store.js) instead of the flat
  // item.purch×stock estimate, which goes stale the moment a purchase is
  // entered at a different rate than what's stored on the item master.
  const fifoMap=useMemo(()=>getFifoStockMap(),[state.purchases.length,state.sales.length,state.repairExpenses.length,state.items.length]);
  const rows=useMemo(()=>{
    const byCat={};
    for(const it of state.items){
      if(it.active===false) continue;
      const cat=it.cat||'Other';
      if(!byCat[cat]) byCat[cat]={cat,count:0,stock:0,stockValue:0};
      const b=byCat[cat];
      const stock=it.stock||0;
      const fifo=fifoMap[it.id];
      b.count++; b.stock+=stock; b.stockValue+=fifo?fifo.value:(stock*(it.purch||0));
    }
    return Object.values(byCat).sort((a,b)=>b.stockValue-a.stockValue);
  },[state.items.length,fifoMap]);
  const totals=useMemo(()=>rows.reduce((s,r)=>({count:s.count+r.count,stock:s.stock+r.stock,stockValue:s.stockValue+r.stockValue}),{count:0,stock:0,stockValue:0}),[rows]);
  const mobile=useIsMobile();
  const columns=[
    {key:'cat',label:'Category',render:r=><Badge v={r.cat} color="blue"/>},
    {key:'count',label:'Items',align:'right',render:r=><span style={{color:C.sub}}>{r.count}</span>},
    {key:'stock',label:'Stock (pcs)',align:'right',render:r=><span style={{fontWeight:600,color:C.txt}}>{r.stock}</span>},
    {key:'stockValue',label:'Stock Value',align:'right',render:r=><span style={{fontWeight:600,color:C.txt}}>{fmt(r.stockValue)}</span>},
  ];
  const footer=rows.length>0&&<div style={{display:'flex',justifyContent:'space-between',padding:mobile?'10px 14px':'9px 12px',
    background:'#F6F8FA',fontWeight:700,fontSize:13,borderTop:mobile?`1px solid ${C.bdr}`:'none',borderRadius:mobile?9:0}}>
    <span>Total</span>
    <span style={{display:'flex',gap:16}}><span>{totals.count}</span><span>{totals.stock}</span><span>{fmt(totals.stockValue)}</span></span>
  </div>;
  return <Modal title="Inventory — Category Breakdown" onClose={onClose} width={560}>
    {catDetail&&<CategoryDetail cat={catDetail} onClose={()=>setCatDetail(null)}/>}
    <div style={{fontSize:12,color:C.sub,marginBottom:12}}>Click a category for item-level detail</div>
    <div style={{border:mobile?'none':`1px solid ${C.bdr}`,borderRadius:8,overflow:'hidden'}}>
      <CardTable mobile={mobile} columns={columns} rows={rows} keyFn={(r,i)=>i}
        onRowClick={r=>setCatDetail(r.cat)} emptyLabel="No stock" footer={footer}/>
    </div>
  </Modal>;
}

function BalanceSheetReport(){
  const bs=useMemo(()=>getBalanceSheet(),[state.txns.length,state.sales.length,state.purchases.length,state.items.length]);
  const[showStock,setShowStock]=useState(false);
  const Line=({label,val,bold,color,onClick})=><div onClick={onClick}
    style={{display:'flex',justifyContent:'space-between',padding:'7px 0',fontWeight:bold?800:500,fontSize:bold?14:13,cursor:onClick?'pointer':'default'}}>
    <span style={{color:C.txt,textDecoration:onClick?'underline dotted':'none',textUnderlineOffset:3}}>{label}</span><span style={{color:color||C.txt,fontVariantNumeric:'tabular-nums'}}>{fmt(val)}</span>
  </div>;
  return <div>
    {showStock&&<StockBreakdownModal onClose={()=>setShowStock(false)}/>}
    <SectionHead title="Balance Sheet" sub="As of today — a snapshot, not a historical statement"/>
    <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16}}>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 22px'}}>
        <div style={{fontSize:12,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:10}}>Assets</div>
        <Line label="Cash & Bank" val={bs.cash}/>
        <Line label="Inventory (stock × cost)" val={bs.inventoryValue} onClick={()=>setShowStock(true)}/>
        <Line label="Accounts Receivable" val={bs.receivables}/>
        <div style={{borderTop:`1px solid ${C.bdr}`,marginTop:8,paddingTop:8}}><Line label="Total Assets" val={bs.assets} bold color={C.blue}/></div>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 22px'}}>
        <div style={{fontSize:12,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:10}}>Liabilities & Equity</div>
        <Line label="Accounts Payable" val={bs.payables} color={C.red}/>
        <div style={{borderTop:`1px solid ${C.bdr}`,marginTop:8,paddingTop:8}}><Line label="Owner's Equity (computed)" val={bs.equity} bold color={bs.equity>=0?C.green:C.red}/></div>
        <div style={{borderTop:`1px solid ${C.bdr}`,marginTop:8,paddingTop:8}}><Line label="Total Liabilities + Equity" val={bs.payables+bs.equity} bold color={C.blue}/></div>
      </div>
    </div>
    <div style={{fontSize:11,color:C.muted,marginTop:10}}>Equity is a residual (Assets − Liabilities), not drawn from a separate owner's-capital ledger.</div>
  </div>;
}

function ReportsModule(){
  const[tab,setTab]=useState('expenses');
  return <div>
    <div style={{display:'flex',gap:8,marginBottom:16}}>
      <Chip v="Expenses" active={tab==='expenses'} onClick={()=>setTab('expenses')}/>
      <Chip v="P&L" active={tab==='pnl'} onClick={()=>setTab('pnl')}/>
      <Chip v="Balance Sheet" active={tab==='balance'} onClick={()=>setTab('balance')}/>
    </div>
    {tab==='expenses'&&<ExpensesReport/>}
    {tab==='pnl'&&<PnLReport/>}
    {tab==='balance'&&<BalanceSheetReport/>}
  </div>;
}

// ── BUSINESS INTELLIGENCE ─────────────────────────────────────────────────────
const monthLabel=(m,full)=>{
  const[y,mo]=m.split('-');
  return new Date(Number(y),Number(mo)-1,1).toLocaleDateString('en-IN',{month:full?'long':'short',year:'2-digit'});
};

function TrendChart({data,series,height}){
  const[hover,setHover]=useState(null);
  const W=760,H=height||280,padL=56,padR=16,padT=16,padB=28;
  const plotW=W-padL-padR, plotH=H-padT-padB;
  const allVals=data.flatMap(d=>series.map(s=>d[s.key]));
  const maxV=Math.max(1,...allVals);
  const minV=Math.min(0,...allVals);
  const range=(maxV-minV)||1;
  const xFor=i=>padL+(data.length>1?(i/(data.length-1))*plotW:plotW/2);
  const yFor=v=>padT+plotH-((v-minV)/range)*plotH;
  const ticks=Array.from({length:5},(_,i)=>minV+range*i/4);
  const onMove=e=>{
    const rect=e.currentTarget.getBoundingClientRect();
    const relX=((e.clientX-rect.left)/rect.width)*W;
    const rel=(relX-padL)/plotW;
    const i=Math.round(rel*(data.length-1));
    setHover(Math.max(0,Math.min(data.length-1,i)));
  };
  const last=data[data.length-1];
  const showEvery=Math.max(1,Math.ceil(data.length/6));
  return <div style={{position:'relative'}}>
    <svg width="100%" viewBox={`0 0 ${W} ${H}`} onMouseMove={onMove} onMouseLeave={()=>setHover(null)} style={{overflow:'visible',display:'block'}}>
      {ticks.map((t,i)=><g key={i}>
        <line x1={padL} x2={W-padR} y1={yFor(t)} y2={yFor(t)} stroke={C.bdr} strokeWidth={1}/>
        <text x={padL-8} y={yFor(t)+3} textAnchor="end" fontSize={9} fill={C.muted}>{fmtL(Math.round(t))}</text>
      </g>)}
      {data.map((d,i)=>(i%showEvery===0||i===data.length-1)&&
        <text key={i} x={xFor(i)} y={H-8} textAnchor="middle" fontSize={9} fill={C.muted}>{monthLabel(d.month)}</text>
      )}
      {series.map(s=>{
        const pts=data.map((d,i)=>`${xFor(i)},${yFor(d[s.key])}`).join(' ');
        return <g key={s.key}>
          <polyline points={pts} fill="none" stroke={s.color} strokeWidth={2} strokeLinejoin="round" strokeLinecap="round"/>
          <circle cx={xFor(data.length-1)} cy={yFor(last[s.key])} r={4} fill={s.color} stroke={C.surf} strokeWidth={2}/>
        </g>;
      })}
      {hover!=null&&<line x1={xFor(hover)} x2={xFor(hover)} y1={padT} y2={padT+plotH} stroke={C.bdrStrong} strokeWidth={1}/>}
      {hover!=null&&series.map(s=><circle key={s.key} cx={xFor(hover)} cy={yFor(data[hover][s.key])} r={4} fill={s.color} stroke={C.surf} strokeWidth={2}/>)}
    </svg>
    {hover!=null&&<div style={{position:'absolute',left:`${(xFor(hover)/W)*100}%`,top:8,transform:`translateX(${hover>data.length/2?'-100%':'0'})`,background:C.txt,color:'#fff',borderRadius:6,padding:'7px 11px',fontSize:11,pointerEvents:'none',whiteSpace:'nowrap',zIndex:10,marginLeft:hover>data.length/2?-10:10}}>
      <div style={{fontWeight:700,marginBottom:4,opacity:.9}}>{monthLabel(data[hover].month,true)}</div>
      {series.map(s=><div key={s.key} style={{display:'flex',alignItems:'center',gap:6}}>
        <span style={{width:9,height:2,background:s.color,display:'inline-block'}}/>
        <span style={{opacity:.75,minWidth:56}}>{s.name}</span>
        <span style={{fontWeight:700}}>{fmt(data[hover][s.key])}</span>
      </div>)}
    </div>}
    {series.length>1&&<div style={{display:'flex',gap:16,marginTop:10,fontSize:11,color:C.sub}}>
      {series.map(s=><div key={s.key} style={{display:'flex',alignItems:'center',gap:6}}>
        <span style={{width:14,height:2,background:s.color,display:'inline-block'}}/>{s.name}
      </div>)}
    </div>}
  </div>;
}

function HBarChart({rows,color,valueFmt}){
  const[hoverI,setHoverI]=useState(null);
  const max=Math.max(1,...rows.map(r=>Math.abs(r.value)));
  if(!rows.length)return <div style={{padding:'24px 0',textAlign:'center',color:C.muted,fontSize:12}}>No data yet</div>;
  return <div>
    {rows.map((r,i)=>{
      const pct=Math.max(2,(Math.abs(r.value)/max)*100);
      return <div key={i} onMouseEnter={()=>setHoverI(i)} onMouseLeave={()=>setHoverI(null)} style={{marginBottom:9}}>
        <div style={{display:'flex',justifyContent:'space-between',fontSize:11,marginBottom:3,gap:8}}>
          <span style={{fontWeight:600,color:C.txt,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.label}</span>
          <span style={{fontWeight:700,color:C.txt,flexShrink:0}}>{valueFmt(r.value)}</span>
        </div>
        <div style={{height:16,background:'#F1F3F5',borderRadius:4,overflow:'hidden'}}>
          <div style={{width:pct+'%',height:'100%',background:color,borderRadius:4,opacity:hoverI===i?0.8:1,transition:'opacity .1s'}}/>
        </div>
      </div>;
    })}
  </div>;
}

function BusinessIntelligence(){
  const monthly=useMemo(()=>getMonthlyBI(12),[state.sales.length,state.purchases.length,state.txns.length,state.items.length]);
  const byCat=useMemo(()=>getMonthlyBIByCategory(12),[state.sales.length,state.payments.length,state.items.length]);
  const monthlyByKey=useMemo(()=>new Map(monthly.map(m=>[m.month,m])),[monthly]);
  const byCatChart=useMemo(()=>byCat.map(m=>({month:m.month,homeSales:m.home.sales,lighterSales:m.lighter.sales,homeProfit:m.home.profit,lighterProfit:m.lighter.profit,homeReceipts:m.home.receipts,lighterReceipts:m.lighter.receipts,netProfit:monthlyByKey.get(m.month)?.profit??0})),[byCat,monthlyByKey]);
  const catTotals=useMemo(()=>byCat.reduce((s,m)=>({
    homeSales:s.homeSales+m.home.sales,homeProfit:s.homeProfit+m.home.profit,homeReceipts:s.homeReceipts+m.home.receipts,
    lighterSales:s.lighterSales+m.lighter.sales,lighterProfit:s.lighterProfit+m.lighter.profit,lighterReceipts:s.lighterReceipts+m.lighter.receipts,
  }),{homeSales:0,homeProfit:0,homeReceipts:0,lighterSales:0,lighterProfit:0,lighterReceipts:0}),[byCat]);
  const netProfitTotal=useMemo(()=>monthly.reduce((s,m)=>s+m.profit,0),[monthly]);
  const byFullCat=useMemo(()=>getSalesProfitByCategory(12),[state.sales.length,state.items.length]);
  const itemMetrics=useMemo(()=>getAllItemMetrics(),[state.sales.length,state.purchases.length,state.items.length]);
  const topSelling=useMemo(()=>[...itemMetrics].filter(m=>m.totalSoldQty>0).sort((a,b)=>b.totalSoldQty-a.totalSoldQty).slice(0,10)
    .map(m=>({label:m.item.internal||m.item.name,value:m.totalSoldQty})),[itemMetrics]);
  const topProfitable=useMemo(()=>[...itemMetrics].filter(m=>m.profit>0).sort((a,b)=>b.profit-a.profit).slice(0,10)
    .map(m=>({label:m.item.internal||m.item.name,value:Math.round(m.profit)})),[itemMetrics]);
  const totals=useMemo(()=>monthly.reduce((s,m)=>({sales:s.sales+m.sales,expenses:s.expenses+m.expenses,profit:s.profit+m.profit}),{sales:0,expenses:0,profit:0}),[monthly]);
  return <div>
    <SectionHead title="Business Intelligence" sub="Last 12 months of sales, expenses & profit · top products by volume and profit"/>
    <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:10,marginBottom:16}}>
      <StatCard label="Sales (12mo)" value={fmtL(totals.sales)} color={C.blue}/>
      <StatCard label="Expenses (12mo)" value={fmtL(totals.expenses)} color={C.amber}/>
      <StatCard label="Profit (12mo)" value={fmtL(totals.profit)} color={totals.profit>=0?C.green:C.red}/>
    </div>
    <div style={{fontSize:11,color:C.sub,marginBottom:10}}>Profit uses current item rates as a cost proxy for older months — a close approximation, not an exact historical figure. Sales is charted separately from Expenses/Profit since it runs roughly 10x larger — one shared axis would flatten the smaller two.</div>
    <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16,marginBottom:16}}>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 20px'}}>
        <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:10}}>Monthly Sales</div>
        <TrendChart data={monthly} series={[{key:'sales',name:'Sales',color:C.blue}]} height={220}/>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 20px'}}>
        <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:10}}>Monthly Expenses & Profit</div>
        <TrendChart data={monthly} series={[{key:'expenses',name:'Expenses',color:C.amber},{key:'profit',name:'Profit',color:C.green}]} height={220}/>
      </div>
    </div>
    <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16,marginBottom:24}}>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 20px'}}>
        <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:2}}>Top Selling Products</div>
        <div style={{fontSize:11,color:C.sub,marginBottom:12}}>By quantity sold, all-time</div>
        <HBarChart rows={topSelling} color={C.blue} valueFmt={v=>v+' pcs'}/>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 20px'}}>
        <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:2}}>Most Profitable Products</div>
        <div style={{fontSize:11,color:C.sub,marginBottom:12}}>By profit earned, all-time</div>
        <HBarChart rows={topProfitable} color={C.green} valueFmt={v=>fmt(v)}/>
      </div>
    </div>

    <div style={{fontSize:16,fontWeight:800,color:C.txt,marginBottom:4}}>Home Appliances vs Lighter</div>
    <div style={{fontSize:11,color:C.sub,marginBottom:14}}>
      Receipts (cash collected) can't be split as cleanly as Sales/Profit — a payment is recorded against a party, not a specific invoice item —
      so this FIFO-matches each customer's dated payments against their dated bills (oldest bill first) and attributes each payment to whichever
      bill(s) it closes. Any leftover past a customer's known bill history (an advance, or a pre-tracked balance) is folded into Home Appliances.
    </div>
    <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:10,marginBottom:16}}>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'14px 16px'}}>
        <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>Sales (12mo)</div>
        <div style={{display:'flex',justifyContent:'space-between',fontSize:13,marginBottom:4}}><span style={{color:C.blue,fontWeight:600}}>Home Appliances</span><b>{fmtL(catTotals.homeSales)}</b></div>
        <div style={{display:'flex',justifyContent:'space-between',fontSize:13}}><span style={{color:C.purple,fontWeight:600}}>Lighter</span><b>{fmtL(catTotals.lighterSales)}</b></div>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'14px 16px'}}>
        <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>Gross Profit (12mo)</div>
        <div style={{display:'flex',justifyContent:'space-between',fontSize:13,marginBottom:4}}><span style={{color:C.blue,fontWeight:600}}>Home Appliances</span><b>{fmtL(catTotals.homeProfit)}</b></div>
        <div style={{display:'flex',justifyContent:'space-between',fontSize:13}}><span style={{color:C.purple,fontWeight:600}}>Lighter</span><b>{fmtL(catTotals.lighterProfit)}</b></div>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'14px 16px'}}>
        <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>Receipts (12mo)</div>
        <div style={{display:'flex',justifyContent:'space-between',fontSize:13,marginBottom:4}}><span style={{color:C.blue,fontWeight:600}}>Home Appliances</span><b>{fmtL(catTotals.homeReceipts)}</b></div>
        <div style={{display:'flex',justifyContent:'space-between',fontSize:13}}><span style={{color:C.purple,fontWeight:600}}>Lighter</span><b>{fmtL(catTotals.lighterReceipts)}</b></div>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'14px 16px'}}>
        <div style={{fontSize:10,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>Net Profit (12mo)</div>
        <div style={{fontSize:20,fontWeight:800,color:netProfitTotal>=0?C.green:C.red}}>{fmtL(netProfitTotal)}</div>
        <div style={{fontSize:11,color:C.muted,marginTop:2}}>Combined — not split by category</div>
      </div>
    </div>
    <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16}}>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 20px'}}>
        <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:10}}>Monthly Sales</div>
        <TrendChart data={byCatChart} series={[{key:'homeSales',name:'Home Appliances',color:C.blue},{key:'lighterSales',name:'Lighter',color:C.purple}]} height={220}/>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 20px'}}>
        <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:10}}>Monthly Gross Profit</div>
        <TrendChart data={byCatChart} series={[{key:'homeProfit',name:'Home Appliances',color:C.blue},{key:'lighterProfit',name:'Lighter',color:C.purple}]} height={220}/>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 20px'}}>
        <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:10}}>Monthly Receipts</div>
        <TrendChart data={byCatChart} series={[{key:'homeReceipts',name:'Home Appliances',color:C.blue},{key:'lighterReceipts',name:'Lighter',color:C.purple}]} height={220}/>
      </div>
      <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,padding:'18px 20px'}}>
        <div style={{fontSize:14,fontWeight:700,color:C.txt,marginBottom:2}}>Monthly Net Profit</div>
        <div style={{fontSize:11,color:C.sub,marginBottom:8}}>Combined (Sales − COGS − business expenses) — not split by category</div>
        <TrendChart data={byCatChart} series={[{key:'netProfit',name:'Net Profit',color:C.green}]} height={196}/>
      </div>
    </div>

    <div style={{fontSize:16,fontWeight:800,color:C.txt,margin:'28px 0 4px'}}>Sales & Profit by Item Category (12mo)</div>
    <div style={{fontSize:11,color:C.sub,marginBottom:14}}>Every real item category, not just the Home/Lighter split above. Receipts aren't broken out this finely — see note above.</div>
    <div style={{background:C.surf,borderRadius:8,border:`1px solid ${C.bdr}`,overflow:'hidden'}}>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
        <thead><tr style={{background:'#F6F8FA'}}>
          {['Category','Qty Sold','Sales','Gross Profit','Margin'].map((h,i)=><th key={i} style={{padding:'9px 12px',textAlign:i>0?'right':'left',fontSize:11,fontWeight:700,color:C.sub,textTransform:'uppercase',letterSpacing:.5,borderBottom:`1px solid ${C.bdr}`}}>{h}</th>)}
        </tr></thead>
        <tbody>
          {byFullCat.length===0&&<tr><td colSpan={5} style={{padding:'30px',textAlign:'center',color:C.muted}}>No sales in the last 12 months</td></tr>}
          {byFullCat.map((c,i)=><tr key={i} style={{borderBottom:`1px solid ${C.bdr}`}}>
            <td style={{padding:'9px 12px',fontWeight:600,color:C.txt}}>{c.cat}</td>
            <td style={{padding:'9px 12px',textAlign:'right',color:C.sub}}>{c.qty} pcs</td>
            <td style={{padding:'9px 12px',textAlign:'right',fontWeight:600,color:C.blue}}>{fmt(c.sales)}</td>
            <td style={{padding:'9px 12px',textAlign:'right',color:c.profit>=0?C.green:C.red}}>{fmt(c.profit)}</td>
            <td style={{padding:'9px 12px',textAlign:'right'}}><Badge v={c.margin+'%'} color={c.margin>25?'green':c.margin>15?'amber':'red'}/></td>
          </tr>)}
        </tbody>
      </table>
    </div>
  </div>;
}

// ── LOGIN (cloud sync gate) ─────────────────────────────────────────────────────
// Only rendered when a Supabase project is configured for this build AND no
// session is cached yet — a dev checkout with no VITE_SUPABASE_* env vars
// set skips this entirely and runs purely local, same as before this
// feature existed. One sign-in per device/browser; session persists after
// that (Supabase's client stores it in localStorage itself).
function LoginScreen({ onSignedIn }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState('');

  const submit = async (e) => {
    e.preventDefault();
    if (!email || !password) { setErr('Enter email and password'); return; }
    setBusy(true); setErr('');
    try {
      await DB.signIn(email.trim(), password);
      onSignedIn();
    } catch (e) {
      setErr(e.message || 'Sign in failed');
    } finally {
      setBusy(false);
    }
  };

  return <div style={{display:'flex',alignItems:'center',justifyContent:'center',height:'100vh',background:C.bg}}>
    <form onSubmit={submit} style={{width:320,background:C.surf,borderRadius:12,border:`1px solid ${C.bdr}`,padding:28,boxShadow:'0 4px 20px rgba(0,0,0,0.06)'}}>
      <div style={{textAlign:'center',marginBottom:20}}>
        <div style={{fontSize:32,marginBottom:8}}>🔥</div>
        <div style={{fontSize:15,fontWeight:700,color:C.txt}}>Upkaran ERP</div>
        <div style={{fontSize:12,color:C.sub,marginTop:4}}>Sign in to sync your data on this device</div>
      </div>
      <div style={{marginBottom:12}}>
        <label style={{display:'block',fontSize:12,fontWeight:600,color:C.sub,marginBottom:4}}>Email</label>
        <input type="email" value={email} onChange={e=>setEmail(e.target.value)} autoComplete="username" autoFocus
          style={{width:'100%',padding:'9px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}/>
      </div>
      <div style={{marginBottom:16}}>
        <label style={{display:'block',fontSize:12,fontWeight:600,color:C.sub,marginBottom:4}}>Password</label>
        <input type="password" value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password"
          style={{width:'100%',padding:'9px 12px',borderRadius:7,border:`1.5px solid ${C.bdr}`,fontSize:13,outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}/>
      </div>
      {err && <div style={{fontSize:12,color:C.red,marginBottom:12}}>{err}</div>}
      <button type="submit" disabled={busy} style={{width:'100%',padding:'10px 0',borderRadius:7,border:'none',background:C.blue,color:'#fff',fontSize:13,fontWeight:700,cursor:busy?'default':'pointer',opacity:busy?0.7:1,fontFamily:'inherit'}}>
        {busy ? 'Signing in…' : 'Sign in'}
      </button>
    </form>
  </div>;
}

// ── APP ───────────────────────────────────────────────────────────────────────
function App(){
  const initialMod=(window.location.hash||'').slice(1);
  const[mod,setMod]=useState(NAV.some(n=>n.id===initialMod)?initialMod:(initialMod?'dashboard':'orders'));
  const[ready,setReady]=useState(false);
  // null = checking cached session, false = show LoginScreen, true = proceed.
  // Stays `true` immediately when this build has no Supabase project
  // configured (supabase===null) — cloud sync is opt-in, a plain checkout
  // still runs fully local with no login gate at all.
  const[authed,setAuthed]=useState(!supabase);
  const mobile=useIsMobile();
  // Desktop: true = sidebar expanded (220px) vs collapsed to icon rail (56px).
  // Mobile: true = drawer open vs closed — starts closed so it doesn't cover
  // the screen on load.
  const[sideOpen,setSideOpen]=useState(()=>window.innerWidth>=MOBILE_BREAKPOINT);

  // DB.initSync() pulls this device's latest cross-device snapshot into
  // localStorage BEFORE store.js's loadAll() reads it back out — so a
  // brand-new device sees the real synced data on its very first load
  // instead of seed data. No-op when cloud sync isn't configured.
  const bootAfterAuth=()=>{ DB.initSync().then(loadAll).then(()=>setReady(true)); };

  useEffect(()=>{
    if(!supabase){ bootAfterAuth(); return; }
    DB.getSession().then(session=>{
      if(session){setAuthed(true);bootAfterAuth();}
      else setAuthed(false);
    });
    return DB.onAuthChange(session=>{
      if(session){setAuthed(true);bootAfterAuth();}
    });
  },[]);

  // Real #hash routing — lets ctrl/cmd/middle-click on a sidebar link open
  // that module in a new browser tab (which reads the hash on load below),
  // instead of only supporting in-tab SPA navigation.
  useEffect(()=>{
    const onHashChange=()=>{
      const id=(window.location.hash||'').slice(1);
      if(NAV.some(n=>n.id===id))setMod(id);
    };
    window.addEventListener('hashchange',onHashChange);
    return()=>window.removeEventListener('hashchange',onHashChange);
  },[]);

  // Cross-tab live refresh — a change saved in another tab (now easy to open
  // via the hash-routed nav links above) pulls its state slice back in via
  // store.js's storage-event listener; remounting the active module here
  // picks it up without the user having to hit F5.
  const[liveTick,setLiveTick]=useState(0);
  useEffect(()=>subscribeLive(()=>setLiveTick(t=>t+1)),[]);

  if(authed===false)return <LoginScreen onSignedIn={()=>{setAuthed(true);bootAfterAuth();}}/>;

  if(!ready)return <div style={{display:'flex',alignItems:'center',justifyContent:'center',height:'100vh',background:C.bg}}>
    <div style={{textAlign:'center'}}>
      <div style={{fontSize:32,marginBottom:8}}>🔥</div>
      <div style={{fontSize:14,color:C.sub}}>Loading Upkaran ERP…</div>
    </div>
  </div>;

  const nav=id=>{window.location.hash=id;setMod(id);if(mobile)setSideOpen(false);};

  const renderMod=()=>{
    switch(mod){
      case'dashboard': return <Dashboard navigate={nav}/>;
      case'sale':      return <SaleEntry/>;
      case'orders':    return <OrdersModule/>;
      case'outstanding':return <Outstanding/>;
      case'parties':   return <PartyMasters/>;
      case'items':     return <ItemMasters/>;
      case'accounts':   return <AccountsModule/>;
      case'reports':    return <ReportsModule/>;
      case'insights':   return <BusinessIntelligence/>;
      case'sales-log':  return <SalesLog/>;
      case'deliveries': return <DeliveriesModule/>;
      case'purch-log':  return <PurchLog/>;
      case'settings':   return <Settings/>;
      default:         return <Dashboard navigate={nav}/>;
    }
  };

  const hasDueFU=state.followups.some(f=>f.status==='open'&&f.followUpDate&&f.followUpDate<=today());

  return <div style={{display:'flex',height:'100vh',background:C.bg,fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',overflow:'hidden'}}>
    {/* Sidebar */}
    {mobile&&sideOpen&&(
      <div onClick={()=>setSideOpen(false)} style={{position:'fixed',inset:0,
        background:'rgba(0,0,0,0.45)',zIndex:199}}/>
    )}
    {(!mobile||sideOpen)&&(
      <div style={{width:mobile?220:(sideOpen?220:56),background:'#0A0E1A',display:'flex',flexDirection:'column',
        flexShrink:0,transition:mobile?'none':'width .18s ease',overflow:'hidden',borderRight:'1px solid #1a2035',
        ...(mobile?{position:'fixed',top:0,left:0,height:'100vh',zIndex:200,boxShadow:'4px 0 24px rgba(0,0,0,0.3)'}:{})}}>
        {/* Logo */}
        <div style={{height:52,display:'flex',alignItems:'center',padding:'0 13px',
          gap:10,justifyContent:sideOpen?'flex-start':'center',borderBottom:'1px solid #1a2035',flexShrink:0}}>
          <div style={{width:30,height:30,borderRadius:8,
            background:'linear-gradient(135deg,#2563EB,#1D4ED8)',
            display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0,cursor:'pointer'}}
            onClick={()=>setSideOpen(o=>!o)}>
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5" strokeLinecap="round">
              <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/>
            </svg>
          </div>
          {sideOpen&&<div>
            <div style={{fontSize:14,fontWeight:800,color:'#FFFFFF',letterSpacing:-.4,whiteSpace:'nowrap'}}>Upkaran</div>
            <div style={{fontSize:9,color:'rgba(255,255,255,.28)',whiteSpace:'nowrap',letterSpacing:.5}}>ERP v{APP_VERSION}</div>
          </div>}
        </div>
        {/* Nav */}
        <nav style={{padding:'8px 6px',flex:1,overflowY:'auto'}}>
          {NAV.map(n=>{
            const active=mod===n.id;
            const dot=n.id==='outstanding'&&hasDueFU;
            return <a key={n.id} href={'#'+n.id} onClick={()=>{if(mobile)setSideOpen(false);}} style={{
              width:'100%',display:'flex',alignItems:'center',textDecoration:'none',boxSizing:'border-box',
              gap:sideOpen?11:0,justifyContent:sideOpen?'flex-start':'center',
              padding:sideOpen?'9px 11px':'9px 0',
              borderRadius:7,border:'none',marginBottom:1,cursor:'pointer',fontFamily:'inherit',
              fontSize:13,fontWeight:active?600:400,
              background:active?'rgba(37,99,235,0.2)':'transparent',
              color:active?'#60A5FA':'rgba(255,255,255,.38)',
              boxShadow:active?'inset 3px 0 0 #2563EB':'none',
              transition:'background .12s,color .12s'}}>
              <span style={{width:20,display:'flex',alignItems:'center',justifyContent:'center',
                flexShrink:0,position:'relative',
                color:active?'#60A5FA':'rgba(255,255,255,.4)'}}>
                {n.icon}
                {dot&&<span style={{position:'absolute',top:-4,right:-5,width:7,height:7,
                  borderRadius:'50%',background:'#F59E0B',border:'2px solid #0A0E1A'}}/>}
              </span>
              {sideOpen&&<span style={{whiteSpace:'nowrap'}}>{n.label}</span>}
            </a>;
          })}
        </nav>
        {/* Footer */}
        <div style={{borderTop:'1px solid #1a2035',padding:'10px 10px',
          display:'flex',alignItems:'center',gap:8,flexShrink:0,
          justifyContent:sideOpen?'space-between':'center'}}>
          {sideOpen&&<div style={{minWidth:0}}>
            <div style={{fontSize:11,fontWeight:600,color:'rgba(255,255,255,.6)',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>Pratik Agrawal</div>
            <div style={{fontSize:10,color:'rgba(255,255,255,.25)'}}>Raipur · CG</div>
          </div>}
          <button onClick={()=>setSideOpen(o=>!o)} title={sideOpen?'Collapse':'Expand'}
            style={{background:'rgba(255,255,255,.07)',border:'none',borderRadius:6,
              cursor:'pointer',padding:'5px 7px',display:'flex',alignItems:'center',
              color:'rgba(255,255,255,.4)',flexShrink:0}}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              {sideOpen?<path d="M15 18l-6-6 6-6"/>:<path d="M9 18l6-6-6-6"/>}
            </svg>
          </button>
        </div>
      </div>
    )}

    {/* Main */}
    <div style={{flex:1,display:'flex',flexDirection:'column',overflow:'hidden'}}>
      {/* Top bar */}
      <div style={{padding:mobile?'0 14px':'0 24px',height:52,display:'flex',alignItems:'center',justifyContent:'space-between',background:C.surf,borderBottom:`1px solid ${C.bdr}`,flexShrink:0,gap:8}}>
        <div style={{display:'flex',alignItems:'center',gap:mobile?8:10,minWidth:0}}>
          {mobile&&<button onClick={()=>setSideOpen(o=>!o)} style={{background:'none',border:'none',cursor:'pointer',fontSize:18,color:C.sub,flexShrink:0,padding:0}}>☰</button>}
          <span style={{fontSize:mobile?13:14,fontWeight:700,color:C.txt,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{NAV.find(n=>n.id===mod)?.label||'Dashboard'}</span>
        </div>
        <div style={{display:'flex',gap:8,flexShrink:0}}>
          <button onClick={()=>nav('sale')} style={{padding:mobile?'6px 10px':'6px 14px',borderRadius:7,background:C.blue,color:'#fff',border:'none',cursor:'pointer',fontSize:mobile?11:12,fontWeight:700,fontFamily:'inherit',whiteSpace:'nowrap'}}>{mobile?'+ Sale':'+ New Sale'}</button>
        </div>
      </div>

      {/* Content */}
      <div style={{flex:1,overflowY:'auto',padding:mobile?'16px 14px 80px 14px':20}}>
        <div key={mod+'#'+liveTick}>{renderMod()}</div>
      </div>
    </div>

    {/* Mobile bottom tabs — curated 4-item set (BOTTOM_NAV), not all of NAV */}
    {mobile&&<div style={{position:'fixed',bottom:0,left:0,right:0,background:C.surf,borderTop:`1px solid ${C.bdr}`,boxShadow:'0 -2px 10px rgba(0,0,0,.05)',display:'flex',padding:'6px 0 calc(4px + env(safe-area-inset-bottom))',zIndex:100}}>
      {BOTTOM_NAV.map(n=>{
        const active=mod===n.id;
        return <a key={n.id} href={'#'+n.id} style={{flex:1,display:'flex',flexDirection:'column',alignItems:'center',gap:3,padding:'2px 0',background:'none',border:'none',textDecoration:'none',cursor:'pointer',fontFamily:'inherit',position:'relative'}}>
          <span style={{display:'flex',alignItems:'center',justifyContent:'center',width:40,height:26,borderRadius:14,background:active?C.blue+'15':'transparent',color:active?C.blue:C.muted,transition:'background .15s'}}>{n.icon}</span>
          <span style={{fontSize:10,fontWeight:active?700:500,color:active?C.blue:C.muted}}>{n.label}</span>
          {n.id==='outstanding'&&hasDueFU&&<span style={{position:'absolute',top:0,right:'28%',width:7,height:7,borderRadius:'50%',background:C.amber,border:`1.5px solid ${C.surf}`}}/>}
        </a>;
      })}
    </div>}
  </div>;
}

export default App;
