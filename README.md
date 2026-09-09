# Firefighter Leave Planner

Leave planning for a fire station rota, built in two generations while I was serving as a firefighter with the Singapore Civil Defence Force.

1. **A Telegram bot** (2025): an LLM agent on n8n with Google Sheets tools. Used by 41 firefighters, 100% adoption, 400+ requests handled with zero double-bookings.
2. **A spreadsheet engine** (2026): the leave rules implemented as Excel VBA and as Google Apps Script, with a Python generator for the workbook.

## The problem

A fire station rota works a fixed shift cycle, each duty day has a limited number of leave slots in priority order, a shift off costs more than one leave day, and leave comes from several pools with different expiry dates. Before this, the calendar lived in a WhatsApp message that an officer re-typed into a spreadsheet, and everyone asked the same three questions: how many days do I have left, is that date free, and who needs to clear leave before they finish service.

## Generation 1: the Telegram bot (`telegram-bot/`)

Three n8n workflows, exported as JSON.

```
Telegram group ──▶ main.json ──▶ OpenAI agent with tools:
                                   • Get Individual Leave Data   (Sheets)
                                   • Get Rota Leave Data         (Sheets)
                                   • Read Calendar               (Sheets)
                                   • Get Personnel Names         (Sheets)
                                   • Enter WhatsApp Calendar     ──▶ whatsapp_parser.json
                                   • Make Calendar Changes       ──▶ make_calendar_changes.json
                   schedule trigger (8 pm before every shift) ──▶ alert listing who must clear leave
```

- **Natural-language queries.** "How many shifts do I have left?", "Who has the highest difficulty score?", "Is the top slot free on 12 July?" The system prompt encodes the domain rules: leave days per shift, slot priority, which slots certain leave types may use, and when a swap needs an officer's confirmation.
- **The agent proposes, deterministic code applies.** When a change is needed the agent returns structured JSON (date, slot, platoon, name). A separate workflow normalises dates, compares the requested state against the live calendar, bumps people up when a higher slot frees, and writes only the rows that actually changed. The model never writes to the sheet directly.
- **WhatsApp calendar import.** Officers still received the official calendar as a WhatsApp message. A JavaScript cleaner turns the free-text message into `date | slot | name` lines and pads every date to the full set of slots, then a second agent resolves nicknames to official names against the personnel sheet and emits one record per date, platoon and slot. The result is diffed against the calendar before writing.
- **Proactive alerts.** The evening before each shift the bot posts everyone whose "difficulty to clear leave" score (a function of leave remaining, shifts left before their service ends, and slots still available) is above a threshold.

The exports have been scrubbed for publication: sheet IDs, chat IDs, credential IDs, station identifiers and personnel names are replaced with placeholders, and test data is removed. Import them into n8n and point the Google Sheets nodes at your own sheet.

## Generation 2: the spreadsheet engine (`workbook/`)

The bot depended on an n8n server and an OpenAI key. For a station that just wants a file, the same rules were reimplemented twice, once as Excel VBA and once as Google Apps Script, so the officer can pick whichever the station uses.

- `generate_workbook.py` builds the workbook with openpyxl: a Calendar sheet showing only this rota's duty dates with dropdown slot columns, a Personnel sheet with balances and a credits log, a Dashboard, and hidden Config and NameLists sheets.
- `vba/` (six modules) and `google_sheets/` (six scripts) implement the same engine: a Calendar change event handler, credit pools with soonest-expiry-first deduction, automatic reversal when a name is removed, slot bumping, shift-filtered name lists, and a dashboard with a clearing tracker that rates how hard it will be for each person to use their leave before they finish service.
- Station parameters (anchor duty date, entitlements, expiry periods, slot counts) are configuration; the values in this repository are placeholders.

See `workbook/README.md` and `workbook/google_sheets/README.md` for setup.

## Stack

n8n, OpenAI, Telegram Bot API, Google Sheets API · Python (openpyxl) · Excel VBA · Google Apps Script

---

Shahu Wagh · [shahuwagh.com](https://www.shahuwagh.com) · [github.com/Shahu-123](https://github.com/Shahu-123)
