# Game Station Lounge POS — Roadmap

## Current Build (v2.0)

### Working Features
- **Table/Tab system** — 18 tables (Table 1–10, VIP 1–3, Bar 1–5) + custom named tabs
- **Payment tracking** — Cash vs MoMo per tab, running totals in stats bar and bottom bar
- **Google Sheets sync** — prices, sales log, closing stock, daily summary via Cloudflare Worker proxy
- **PIN protection** — manager PIN required for Restock and End Day (lockout after 3 failed attempts)
- **Auto stock loading** — yesterday's closing stock pre-fills opening stock on new day
- **Session persistence** — localStorage keeps session alive across browser refreshes
- **Professional receipts** — printable per-tab HTML receipt with full itemisation
- **End of day report** — closed tabs list, cash/momo breakdown, grand total, PDF export
- **Stock management** — real-time remaining stock, low stock warnings, restock modal
- **Background image** — applied to loading, setup, and app screens

---

## Known Issues / Watch List
- Cloudflare Worker proxy adds ~2–5s latency to each API call
- GitHub Pages CDN can take up to 10 min to serve newly pushed files
- End Day save splits into two parallel API calls (summary + closing stock) — if one fails, data may be partially saved
- `localStorage` cleared on End Day success; if browser is closed mid-day without ending day, session restores correctly on reopen
- WhatsApp low stock alerts removed pending Callmebot API key

---

## Future Features

### High Priority
- **Split payment** — part Cash, part MoMo on the same tab (e.g. 3000 cash + 2000 momo)
- **WhatsApp low stock alerts** — Callmebot API (API key pending); replaces removed email alerts
- **Staff login** — multiple bartenders each with their own PIN, sales attributed per staff member

### Reporting & Analytics
- **Monthly sales report** — auto-generated Google Sheets tab summarising revenue by month
- **Manager dashboard** — read-only web view to monitor live sales and stock remotely without opening the POS
- **Product performance** — most/least sold items over time, revenue per category

### Product Experience
- **Product images** — photo on each product card (stored in Google Drive or inline base64)
- **Favourites / quick-add** — pin the top 5 most-ordered drinks to a quick-access row

### Customer & Loyalty
- **Customer loyalty tracking** — optional tab name lookup; track visit count and lifetime spend
- **Discount codes** — apply % or fixed discount on a tab before payment

### Stock & Operations
- **Inventory reorder alerts** — configurable threshold per product; alert when stock falls below X
- **Supplier management** — record deliveries with supplier name, date, quantities
- **Waste tracking** — log broken/expired stock separately from sales

### Technical
- **Offline mode** — queue API calls locally when no internet; sync when connection restores
- **Mobile app** — package as PWA (Progressive Web App) with home screen install, push notifications
- **Custom proxy** — replace Cloudflare Worker with a dedicated lightweight server to reduce latency
