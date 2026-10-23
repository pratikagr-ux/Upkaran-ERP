# Upkaran ERP v2.0

Full business ERP for Upkaran kitchenware distributorship, Raipur CG.

## Setup

```bash
npm install
npm run dev
```
Opens at http://localhost:3000

## Project Structure

```
src/
├── main.jsx          Entry point
├── App.jsx           Main shell + sidebar + router
├── db.js             localStorage storage layer
├── store.js          All module state, load/save, rate helpers
├── constants.js      Colors, formatters, account definitions
├── data/             Seed data (parties, items, rates, history)
│   ├── items.js      188 products
│   ├── parties.js    422 parties (411 customers, 11 suppliers)
│   ├── lastRates.js  Party-item last rates
│   ├── modeRates.js  Modal rates per item
│   ├── partyInvoices.js  Outstanding invoice breakdown
│   └── partyHistory.js   Sales & payment history
└── components/
    ├── ui/           Badge, Btn, Inp, Modal, StatCard etc.
    ├── Dashboard.jsx
    ├── Outstanding.jsx  + OutstandingDetail
    ├── Parties.jsx      + PartyProfile (with AI summary)
    ├── SaleEntry.jsx    + InvoiceModal + EditSale + SalesReturn
    ├── SalesLog.jsx
    ├── PurchLog.jsx
    ├── Accounts.jsx     Cash & Accounts module
    └── Settings.jsx     + BackupRestore
```

## Accounts

Opening balances (from Excel cut-off May 2026):
- Cash in Hand: ₹8,650
- MM (Mahamaya Minerals): ₹55,543
- Ashi ICICI (consolidated): −₹2,31,247
- Upkaran A/c: ₹4,287

## Data Persistence

All ERP entries (sales, payments, transactions, follow-ups) are saved to 
localStorage. Use Settings → Backup to download a JSON file periodically.

## Features

- New Sale with keyboard-first entry
- Outstanding with aging buckets + Lighter dues
- Party profiles with sales/payment history + AI summary
- Cash & Accounts with 4-account ledger
- Follow-up tracking from Outstanding responses
- Sales returns with per-invoice discount/write-off
- Backup & Restore (JSON)
