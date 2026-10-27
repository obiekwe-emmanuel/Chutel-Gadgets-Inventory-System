# Chutel Gadgets Inventory System

A full perpetual inventory and sales management system, built entirely in Excel using VBA. No external database, no add-ins, just a single `.xlsm` workbook that behaves like a lightweight ERP.

---

## The Problem

Picture a small gadgets and accessories store in a busy market. Stock split across two or three storage locations because the main shop can't hold it all. Every purchase and every sale tracked by hand, in a notebook.

The real cost isn't the notebook itself, it's the hours lost every week manually calculating what's left: what was bought, minus what's been sold, just to get a rough guess. And a business owner spending that much time on manual admin is a business owner who isn't spending that time growing the business.

This project is a working fix for exactly that problem.

---

## What It Does

- **Add products** through a form, SKU codes generate automatically, reusing category prefixes so numbering never clashes or resets.
- **Log purchases** through a form, date fills in automatically, totals calculate live as you type.
- **Log sales** through a form, with two hard validation stops: a sale is blocked if the price falls outside that product's allowed range, and blocked again if the quantity exceeds what's actually in stock.
- **Instant stock alerts**, the moment a valid sale saves, a custom-styled pop-up shows exactly how many units of that product remain, no calculating required.
- **Automatic profit and loss**, calculated on every sale, color-coded green for profit and red for loss.
- **A live inventory sheet**, stock levels update automatically the instant a purchase or sale happens anywhere else in the system.
- **A KPI dashboard**, total revenue, profit and loss, cost of goods sold, total transactions, average order value, best-selling product, and top customer by spend, all calculated automatically from the underlying data.
- **Code-level sheet protection**, data sheets are locked against manual editing, but the VBA forms can still write to them normally.
- **Autocomplete fields** for product categories and customer names, reducing duplicate or misspelled entries.

---

## Screenshots

### Landing Page
![Landing Page](screenshots/GA Landing Page.png)

### Add New Product Form
![Add New Product Form](screenshots/GA Add New Product Form.png)

### Database Product Sheet
![Database Product Sheet](screenshots/GA Database Product Sheet.png)

### Record Purchase Form
![Record Purchase Form](screenshots/GA Record Purchase Form.png)

### Purchase Saved Pop Up
![Purchase Saved Pop Up](screenshots/GA Purchase Saved Pop up.png)

### Purchases Sheet
![Purchases Sheet](screenshots/GA Purchases Sheet.png)

### Record Sale Form
![Record Sale Form](screenshots/GA Record Sale Form.png)

### Invalid Sale Price Pop Up
![Invalid Sale Price Pop Up](screenshots/GA Invalid Sale price Pop up.png)

### Remaining Stock Pop Up
![Remaining Stock Pop Up](screenshots/GA Remaining Stock Pop up.png)

### Sales Sheet
![Sales Sheet](screenshots/GA Sales Sheet.png)

### Inventory Sheet
![Inventory Sheet](screenshots/GA Inventory Sheet.png)

### Dashboard Sheet
![Dashboard Sheet](screenshots/GA Dashboard Sheet.png)

---

## How It's Built

**5 linked sheets:**

| Sheet | Purpose |
|---|---|
| `Landing_Page` | Home base, branding, instructions, and the buttons that open every form |
| `Database_Products` | The master product catalog, every other sheet pulls from this one |
| `Purchases` | Logs stock coming in from suppliers |
| `Sales` | Logs stock going out to customers, with profit/loss calculated per row |
| `Inventory` | A live, formula-driven view of current stock per product |
| `Dashboard` | Automatically calculated business KPIs |

**VBA UserForms** handle all data entry, `frmAddProduct`, `frmPurchase`, and `frmSale`, so nobody types directly into the data sheets. A custom `frmSaleAlert` form replaces Excel's default message box for the remaining-stock notification, styled to match the project's branding.

**Key logic highlights:**
- SKU generation scans existing products for a matching category prefix and continues its numbering, or builds a new prefix if the category is new.
- The Sale form checks price range before stock availability, catching simple typos before running the heavier inventory check.
- Inventory and Dashboard are pure formula sheets, no VBA writes to them directly, which keeps them reliable as the single source of truth.
- Sheets are protected using `UserInterfaceOnly:=True`, which blocks manual typing while still allowing the VBA forms to write data normally.

---

## Tech Stack

- Microsoft Excel (2021 / Microsoft 365)
- VBA (UserForms, event-driven macros, Scripting.Dictionary for de-duplication)
- Native Excel formulas (SUMIF, INDEX/MATCH, UNIQUE, FILTER)

---

## About This Project

This is a portfolio project, built using a fictional business scenario to demonstrate a real, common small-business pain point: manual, error-prone inventory tracking across multiple storage locations.

Built by **Obiekwe Emmanuel C.**

