# Handover: Jatropha seed storage and sowing plan (Oct–Dec 2026)

These notes are for continuing this work in Claude Code. In a new session, have Claude read this file first.

**Local working folder:** `C:\Users\kazdr\OneDrive - 河村研究室\Malema_2\01_planning`
It has the same content as the `planning/` folder of the GitHub repo `mameyompo-maker/jatropha_record`, on branch `ccr-f2d35b55-0cwtw4`.

The Japanese version of this file is `HANDOVER.md`.

## 1. Purpose

Plan how much seed must be stored (warehouse and drying area) and when each lot is sown. The output is a Word document with:

- a stacked chart of storage by date, with one colour per seed lot
- a summary of how storage rises and falls, the peak, and when there is not enough seed to sow

## 2. Files

| Path | Contents |
|---|---|
| `種子保管・播種計画_2026.docx` | Deliverable (Word, Japanese) |
| `Seed_Storage_Sowing_Plan_2026_EN.docx` | Deliverable (Word, English) |
| `種子保管量グラフ_2026.png` / `Seed_Storage_Chart_2026_EN.png` | Chart on its own (Japanese / English) |
| `scripts/sim.py` | Day-by-day simulation and charts. Writes `plan.json`, `stock.png` (Japanese) and `stock_en.png` (English) |
| `scripts/make.js` | Builds the Japanese Word file (uses the npm package `docx`) |
| `scripts/make_en.js` | Builds the English Word file |
| `scripts/plan.json` | Simulation result: schedule by lot and daily stock |

Paths are relative to the `01_planning` folder (`planning/` in the repo).

## 3. How to rebuild

You need Python and Node.js. Run these from the `01_planning` folder (Windows example):

```powershell
pip install matplotlib            # first time only
npm install docx                  # first time only (creates node_modules in 01_planning)
python scripts\sim.py             # updates plan.json and the chart PNGs, prints the lot schedule
node scripts\make.js              # builds the Japanese docx
node scripts\make_en.js           # builds the English docx
```

Output paths are based on where the scripts are (the folder above `scripts`), so files are written to the same place wherever you run the commands from. Close the Word files first; if a document is open, the overwrite fails.

## 4. Current conditions (confirmed)

**Seed lots** (sown in this order)

| # | Lot | kg | Arrival |
|---|---|---|---|
| 1 | Ribáuè 2026 | 1,326 | In stock (already dry) |
| 2 | Ribáuè 2025 | 749 | In stock (already dry) |
| 3 | Ghana 0 | 745 | In stock (already dry) |
| 4 | India | 8,400 | 2026-10-16 |
| 5 | Ghana 1 | 3,000 | 2026-10-21 |
| 6 | Ghana 2 | 12,000 | 2026-11-14 |
| 7 | Ghana 3 | 3,700 | 2026-11-13 |

Total: 29,920 kg.

**Rules**

- 100 workers sow **560 kg per day in total across all lots**. Saturday is a half day (280 kg); Sunday is off.
- If a lot runs out partway through a day, the rest of that day's amount comes from the next lot.
- Each new arrival dries for 7 days, counting the arrival day. Day 8 is for preparation and soaking, and sowing can start on day 9.
- **First sowing day is 2026-10-21 (Wed).** The lots in stock (#1–3) are prepared and soaked on 10-20.
- Preparation and soaking is lead time for each lot; other lots can still be sown on that day.
- The set order is kept. Ghana 3 is ready before Ghana 2, but it is still sown after Ghana 2.
- Storage means the amount left after each day's sowing, including seed that is drying.
- Public holidays are not included (25 Dec is counted as a working day).

## 5. Current results

| Lot | Sowing period |
|---|---|
| Ribáuè 2026 | 21–23 Oct |
| Ribáuè 2025 | 23–26 Oct |
| Ghana 0 | 26–27 Oct |
| India | 27 Oct – 16 Nov |
| Ghana 1 | 16–21 Nov |
| Ghana 2 | 23 Nov – 18 Dec |
| Ghana 3 | 18–28 Dec |

- **Finish date:** Mon 28 Dec. It is set by the arrival of Ghana 2 and Ghana 3, so changing the start date does not move it. If 25 Dec is a day off, sowing finishes on 29 Dec.
- **Peak storage:** 18,720 kg (14–15 Nov). Plan for about 19 tonnes of space.
- **Days without enough seed:** only 60 kg short on Sat 21 Nov.
- **Seed drying at the same time:** 11,400 kg on 21–22 Oct and 15,700 kg on 14–19 Nov.

## 6. Change history

1. First version: 560 kg every day, with the lots in stock sown from 10-17.
2. Sundays off and half-day Saturdays. The 560 kg/day is the total across all lots.
3. Ghana 0 (745 kg, in stock) added after Ribáuè 2025 and before India.
4. First sowing day moved to 10-21 (current version).
5. English version added (Word and chart).

## 7. Things to know when changing conditions

- **Where the inputs are:** lots, amounts, arrival dates and colours are in `lots` in `sim.py`. The start date is in `sim.py` at `prep=D(2026,10,20); ready=D(2026,10,21)`.
- **Chart settings:** the start of the red shortage bands is `t["date"]>=D(2026,10,21)` in `sim.py`. The dotted line and its label use `axvline`. All chart wording, Japanese and English, is in `TEXT` in `sim.py`.
- **The body text in `make.js` and `make_en.js` is written by hand.** The figures in the key points, the storage-change table (4.1), the shortage table (4.2) and the suggestions (section 5) are not calculated automatically. After rerunning `sim.py`, read its output and `plan.json`, then update the text in **both** languages. These update by themselves: the schedule-by-lot table, the daily appendix table, the charts and the row colours (also check the start date in `shortSet`).
- **More than 7 lots:** colours go up to the 8th colour of the categorical palette (red `#e34948`). Beyond that, group lots together instead of adding colours.
- **Fonts:** the charts use whichever of IPA / Yu Gothic / Meiryo / Hiragino is available (Yu Gothic on Windows). The Japanese document uses Yu Gothic; the English document uses Arial.
- **Page breaks:** there is a page break before section 4 so its heading is not left alone at the bottom of a page.

## 8. Possible next steps (not started)

- An option to make public holidays (such as 25 Dec) days off
- Generating the body text of `make.js` / `make_en.js` from `plan.json`, so it no longer has to be edited by hand after each change
- A chart that shows the drying area and the warehouse (after drying) separately
