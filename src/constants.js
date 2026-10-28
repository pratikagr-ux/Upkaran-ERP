export const APP_VERSION = '2.0';

export const C = {
  bg:'#F3F4F6', surf:'#FFFFFF', bdr:'#E5E7EB', bdrStrong:'#9CA3AF',
  txt:'#111827', sub:'#4B5563', muted:'#9CA3AF',
  blue:'#2563EB', blueBright:'#1D4ED8', blueSoft:'#EFF6FF', blueMid:'#93C5FD',
  amber:'#D97706', amberBright:'#B45309', amberSoft:'#FEF3C7',
  green:'#059669', greenBright:'#047857', greenSoft:'#ECFDF5',
  red:'#DC2626', redBright:'#B91C1C', redSoft:'#FEF2F2',
  purple:'#7C3AED', purpleSoft:'#F5F3FF',
  sidebar:'#0A0E1A', sidebarBorder:'rgba(255,255,255,0.07)',
};

export const fmt  = n => '₹' + Number(n || 0).toLocaleString('en-IN');
export const fmtL = n => n>=10000000?`₹${(n/10000000).toFixed(2)}Cr`:n>=100000?`₹${(n/100000).toFixed(1)}L`:n>=1000?`₹${(n/1000).toFixed(0)}K`:fmt(n);
export const fmtD = d => { try{ return new Date(d+'T00:00:00').toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'2-digit'}); } catch { return d||'—'; }};
export const today = () => new Date().toISOString().slice(0,10);
export const uid   = p => p + Date.now().toString(36).toUpperCase();

export const ROUTES = ['CG North','CG South','Odisha','Raipur'];

export const ACCOUNTS = [
  {id:'cash',    label:'Cash in Hand',  short:'Cash',    ob:8650,    color:'#16A34A', icon:'💵'},
  {id:'mm',      label:'MM (Mahamaya)', short:'MM',      ob:55543,   color:'#2563EB', icon:'🏦'},
  {id:'ashi',    label:'Ashi ICICI',    short:'Ashi',    ob:-231247, color:'#DC2626', icon:'🔴'},
  {id:'upkaran', label:'Upkaran A/c',   short:'Upkaran', ob:4287,    color:'#7C3AED', icon:'🟢'},
];
export const ACC_MAP = Object.fromEntries(ACCOUNTS.map(a => [a.id, a]));

export const EXPENSE_CATS = [
  {id:'transport',   label:'Transport'},
  {id:'office',      label:'Office Expenses'},
  {id:'tour',        label:'Tour Expenses'},
  {id:'salary',      label:'Salary'},
  {id:'petrol',      label:'Petrol'},
  {id:'rent',        label:'Rent'},
  {id:'delivery',    label:'Delivery'},
  {id:'vendor',      label:'Vendor Payment'},
  {id:'transporter_payment', label:'Transporter Payment'},
  {id:'ghar_kharch', label:'Ghar Kharch'},
  {id:'personal',    label:'Personal / Family'},
  {id:'other',       label:'Other'},
];

// Pre-ERP cumulative expense totals per category, pulled from the "Party balances"
// sheet (rows 543-579, bank-entry-punched expense heads). Individual line items
// (e.g. per-driver salaries, per-route transport) are summed into their matching
// category. Used as the historical baseline in the Expenses/P&L report.
export const HISTORICAL_EXPENSES = {
  transport: 301715,   // Auto/Bus Bhada, Paid/Delhi Transportation, Pankaj Logistics
  delivery: 9065,
  office: 583740,
  tour: 359398,
  salary: 478520,       // Durgesh, upk benu, Anil Sahu, Ravi Thakur, Bhawani/Hanumant Driver, Danish
  petrol: 20500,
  rent: 248000,         // Godown Rent + Bhaisthan Godown Rent
  ghar_kharch: 2742658,
  // Self Consumption used to be lumped in here, but it's goods taken by the
  // owner, not a cash expense — it's now recorded as a real sale (party
  // "Self Consumption") so it shows up in Revenue instead.
  personal: 0,
  // Lost/Damaged/Extra/Adjustment/Counter Sales were Sales-sheet placeholder
  // "customer" rows recording what that write-off inventory would have sold
  // for — not a cash outflow. The inventory's actual cost was already spent
  // when it was purchased (reflected in COGS via HISTORICAL_COGS), so
  // expensing this again here would double-count the same cost. Left at 0
  // rather than removed so the category still works for genuine new
  // miscellaneous cash expenses going forward.
  other: 0,
};

// Total historical Cost of Goods Sold, computed transaction-by-transaction
// from the Purchase sheet as Sales qty × Purchase Rate (column K, the raw
// rate — NOT column P "Cost/unit", which blends in an allocated transport
// cost via column N). Transport is already its own line in Expenses, so
// including it here too would double-count it. Using the actual rate paid at
// each purchase is also far more accurate than applying one current rate
// across years of sales (which previously overstated gross profit by ~₹3.8M).
// New in-app sales fall back to a qty × current-raw-rate proxy (item.purchRaw)
// since there's no better per-sale cost data for those yet.
export const HISTORICAL_COGS = 29647103.97;

// Default business/non-business tag per category — user can override in the
// Expenses report. Ghar Kharch and personal draws are non-business by default.
export const DEFAULT_CATEGORY_BIZ = {
  transport:'business', delivery:'business', office:'business', tour:'business',
  salary:'business', petrol:'business', rent:'business', other:'business',
  ghar_kharch:'personal', personal:'personal',
};

// Unified category list for the Accounts entry form — each category carries the
// direction (in/out/contra) money moves by default, and which kind of party (if any)
// it's linked to. Selecting a category auto-sets direction; the form lets the user
// flip it manually for edge cases (e.g. a vendor refund is a Vendor Payment category
// but flows 'in').
export const TXN_CATEGORIES = [
  {id:'party_receipt', label:'Party Payment / Receipt', dir:'in',  party:'customer',    icon:'💰'},
  {id:'other_receipt', label:'Other Receipt',           dir:'in',  party:null,          icon:'↑'},
  {id:'vendor',                label:'Vendor Payment',       dir:'out', party:'supplier',    icon:'🏭'},
  {id:'transporter_payment',   label:'Transporter Payment',  dir:'out', party:'transporter', icon:'🚚'},
  ...EXPENSE_CATS.filter(c=>c.id!=='vendor'&&c.id!=='transporter_payment').map(c=>({...c,dir:'out',party:null,icon:'↓'})),
  {id:'transfer', label:'Transfer', dir:'contra', party:null, icon:'↔'},
];

export const TAG_COLORS = {
  promise: '#2563EB', paid: '#059669', unavailable: '#D97706',
  'no-answer': '#4B5563', dispute: '#DC2626', note: '#7C3AED',
  'sales-followup': '#0D9488',
};
