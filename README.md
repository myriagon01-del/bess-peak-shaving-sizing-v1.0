# BESS Peak-Shaving Sizing Tool

A self-contained, browser-based tool that sizes a **Battery Energy Storage System
(BESS)** to shave a customer's peak electricity demand — reducing demand charges
and peak-load penalties.

Open **`index.html`** in any modern browser. No install, no server, fully offline.

---

## Features

- Upload an interval **load profile** (`.xlsx` or `.csv`).
- Configure the **sizing window** (the daily peak window) and the **demand interval**
  (the interval over which peak demand is measured).
- Three sizing strategies: **Static Threshold**, **Dynamic Discharge**, and
  **Constant Discharge**.
- Sizes for a **reserve margin** and **Year-15 degradation**.
- Built-in **battery / PCS library** (editable, with a downloadable Excel template).
- Live chart overlaying **load**, **shave-to target**, **battery discharge**, and
  **grid after BESS**, plus a peak before/after/reduction summary.

> The bundled battery/PCS models are **illustrative reference examples** named after
> well-known manufacturers. Their specifications are approximate and must be verified
> against current datasheets before any real sizing.

---

## Quick start

1. Open `index.html` in a browser.
2. Upload a load profile (use **"Download CSV template"** for the expected columns,
   or **"Adjust mapping"** to map your own file).
3. Set the sizing window, demand interval, target and battery model.
4. Read the recommended sizing and check the chart.

---

## Project layout

| Path | Purpose |
|---|---|
| `index.html` | The built single-file app — open this. |
| `src/app.js` | Sizing engine + UI logic. |
| `src/body.html` | Layout. |
| `src/style.css` | Theme. |
| `src/xlsx.umd.min.js` | Vendored SheetJS (Excel parsing). |
| `src/chart.umd.min.js` | Vendored Chart.js (4.4.1). |
| `library/bess_library_template.xlsx` | Battery/PCS library template (embedded at build time). |
| `build/build.js` | Rebuilds `index.html` from `src/`. |

---

## Rebuilding

Requires **Node.js ≥ 18**.

```bash
node build/build.js
```

---

## How the sizing works (plain English)

- **Sizing window** — the daily window during which peak demand is billed; only
  demand inside it is shaved.
- **Demand interval** — the interval (minutes) over which peak demand is measured.
- **Governing day** — the day that needs the most energy to shave; it sets the
  battery size.
- **Energy** — `packs = ceil(required energy × (1 + reserve) ÷ usable per pack at Year 15)`.
- **Power & PCS** — sized separately; the tool never silently substitutes one for the other.

---

## Load profile formats

The tool accepts generic interval data in `.xlsx` or `.csv`. The **"Adjust mapping"**
panel lets you map columns (timestamp, load, optional PV generation / grid export),
set the unit (**kW / MW / W**), the interval, and whether timestamps mark the start
or end of an interval. Every field is documented in the **"Parameter Definitions"**
sheet of the library template.

---

## Security note

This build has **no passcode / login gate** — it opens directly with no access
code required.

## License

MIT — see [`LICENSE`](LICENSE). Bundled third-party libraries (Chart.js,
SheetJS) are listed in [`THIRD_PARTY_LICENSES.md`](THIRD_PARTY_LICENSES.md).
