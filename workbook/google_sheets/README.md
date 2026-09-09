# Station Leave Manager — Google Sheets Version

Leave management for one rota of a fire station. Firefighters plot leave directly on a Calendar sheet via dropdowns. Balances, smart deductions, and slot bumping happen automatically.

## Setup (one-time, ~5 minutes)

### 1. Create a new Google Sheet

Go to [sheets.new](https://sheets.new) to create a blank spreadsheet.

### 2. Open the Script Editor

**Extensions > Apps Script**

### 3. Paste the code

You need to create **6 script files**. For each file:

1. In the Apps Script editor, click the **+** next to "Files" (left sidebar)
2. Select **Script**
3. Name the file (delete the default `Untitled` text)
4. Delete any boilerplate code
5. Paste the contents from the corresponding `.gs` file

Create these files in order:

| File Name          | Source File             | Purpose                           |
|--------------------|------------------------|-----------------------------------|
| `Constants`        | `Constants.gs`         | Column indices, config, helpers   |
| `Setup`            | `Setup.gs`             | Creates all sheets & formatting   |
| `CalendarHandler`  | `CalendarHandler.gs`   | Handles Calendar cell edits       |
| `LeaveEngine`      | `LeaveEngine.gs`       | Credit pools, deductions, balance |
| `Dashboard`        | `Dashboard.gs`         | Dashboard refresh routines        |
| `Menu`             | `Menu.gs`              | Custom menu & refresh functions   |

> Delete the default `Code.gs` file after creating the 6 files above.

### 4. Run Setup

1. In the Apps Script editor, select `setupWorkbook` from the function dropdown (top toolbar)
2. Click **Run** (play button)
3. On first run, Google will ask you to **authorize** the script — click through the prompts
4. Wait for the setup to complete (creates 5 sheets, installs trigger)

### 5. Go back to your spreadsheet

You should see:
- **Calendar** — the rota's duty dates with dropdown-enabled slot columns
- **Personnel** — Empty table ready for your people
- **Dashboard** — Officer overview (populated after first Refresh All)
- A custom menu: **Station Leave Manager** (top menu bar)

---

## Adding Personnel

Go to the **Personnel** sheet and fill in rows starting at row 4:

| Column | Field | Example |
|--------|-------|---------|
| A | Name | Ahmad Bin Ismail |
| B | Rank | CPL |
| C | Join Date | 01-Jan-2025 |
| D | ORD Date | 31-Dec-2026 |
| E | Status | Active (dropdown) |
| F | VL Entitlement | per station policy |

After adding people, go to **Station Leave Manager > Refresh All**.

This:
- Populates the Calendar dropdowns with your personnel
- Auto-credits VL for anyone who doesn't have it yet
- Updates all balances and the Dashboard

---

## Plotting Leave

1. Go to the **Calendar** sheet
2. Click any slot cell (L1/L2/KIV columns) on a date row
3. Select a name from the dropdown
4. Done — leave balance is automatically deducted (2 days, soonest-expiry-first)

**To remove:** Delete the cell content. The person's leave is reversed and everyone bumps up in priority.

**To swap:** Select a different name. The old person's leave is reversed, the new person's is deducted.

---

## Crediting PHOL and OIL

### Automatic PHOL

Go to **Station Leave Manager > Credit PHOL for Past PHs**. This automatically credits 1 PHOL day for each past public holiday that fell on a duty date.

### Manual Credits

Add entries directly in the **Leave Credits & Debits Log** section on the Personnel sheet (row 58+):

1. Find the next empty row
2. Enter: ID (next number), Display Name (e.g. "CPL Ahmad"), CREDIT, PHOL or OIL, days, date, expiry date, remarks
3. Run **Refresh All**

Expiry dates follow the values on the Config sheet.

---

## Dashboard

Go to **Leave Manager > Refresh Dashboard** (or Refresh All) to update the key metrics, the leave-clearing tracker, monthly slot usage, expiry watch and utilisation figures.

## Configuration

The hidden Config sheet holds station parameters (name, year, anchor duty date, entitlement, expiry periods, slot count) and the public holiday table. Unhide it to edit, then re-hide.
