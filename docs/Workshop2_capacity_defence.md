# Workshop 2: the capacity figure and the assumptions behind it

**Subject 2, Financial Capacity perimeter.** Working note for the group, 7 October 2026.
Every number here comes from `model/Veolia_capacity_2027.xlsx`. Each input in that workbook cites a document and PDF page. To rebuild the workbook, run `python model/build_capacity_model.py`, then recalculate it in Excel or LibreOffice.

---

## 1. The figure we defend

> **Between 1 July 2026 and 31 December 2027, Veolia can commit about €1.35bn of acquisition enterprise value without ending 2027 above 3.0x net debt / EBITDA. The defensible range is €0.3bn to €2.9bn.**

| Case | EBITDA 2027e | NFD Dec-27 before new M&A | Leverage before M&A | **Capacity** |
|---|---|---|---|---|
| Base (trajectory) | 7,848 | 22,192 | 2.83x | **1,353** |
| GreenUp EBITDA target met (€8.0bn) | 7,910 | 22,192 | 2.81x | **1,538** |
| Downside | 7,788 | 23,046 | 2.96x | **318** |
| Upside (all disposals done, pro forma ratio) | 7,914 | 21,692 | 2.74x | **2,928** |

Figures are in € m. EBITDA in every case is net of the EBITDA lost through disposals.

**Where the base €1.35bn comes from (sources of capacity, block 4):**

| Source | € m |
|---|---|
| Retained FCF, H2-26 + 2027 (net FCF less dividends, hybrid coupons, buyback) | +856 |
| Disposals (€1.5bn cashed) less the debt room they cost through lost EBITDA | +1,230 |
| Leverage room versus June-2026 debt (3.0x × EBITDA 2027 − €24,548m) | **−733** |
| Pro forma uplift from acquired EBITDA | 0 |
| **Capacity** | **1,353** |

**The finding to lead with.** Today's debt sits *above* what 3x of 2027 EBITDA allows. Retained cash flow covers that gap with only about €0.1bn to spare. **So in practice, the acquisition capacity to end-2027 is the disposal programme.** Without disposals, capacity is close to zero. This agrees with the briefing's warning that Veolia "has to reduce debt to keep its promise". It also puts a number on that warning.

**Model validation.** With no tuning, the model gives **3.08x at end-2026**. Veolia's guidance is "3x or slightly above, Clean Earth included" (H1-26 report, p.29). We did not calibrate the model to hit that figure, so the match is a real check on the roll-forward. Use it when the jury asks why they should trust the model.

---

## 2. Block 1: what has already been spent

| | € m |
|---|---|
| Envelope: "€4bn in growth investments, of which €2bn for boosters" | 4,000 |
| Named deals 2024 to H1-26 (Hofmann, Danubius, WTS 30%, US tuck-ins, Zeeklite, Enviropacific, Clean Earth) | 5,333 |
| of which boosters | 5,018 |
| All financial acquisitions 2024 to H1-26, gross | 5,721 |
| All financial acquisitions, net of financial disposals | 4,602 |

**On any reading, the €4bn envelope is already used up.** Gross spending exceeds it by €1.7bn. Even net of disposals, it exceeds it by €0.6bn. **Any further deal is outside the plan as announced.** That is why our capacity number comes from the leverage constraint and not from "what is left of the €4bn".

Two ambiguities to state in the report, not hide:
- **Gross or net.** Veolia's own wording does not say. We show both.
- **Which period.** The briefing (p.20) says "€4bn earmarked for acquisitions over 2023-2027". The GreenUp launch text (H1-24 report, p.12) says "€4bn in growth investments" for 2024-2027. "Growth investments" may also cover discretionary capex. FY2023 is not in our documents, so 2023 is an open gap.

---

## 3. Block 2: starting position

| | Dec-24 | Jun-25 | Dec-25 | Jun-26 |
|---|---|---|---|---|
| Net financial debt | 17,819 | 20,764 | 19,657 | 24,548 |
| EBITDA (FY or last twelve months) | 6,788 | 6,889 | 7,050 | 7,235 |
| Leverage | 2.63x | 3.01x | 2.79x | 3.39x |

- **Definition.** We use Veolia's own: "financial debt including IFRS 16 net closing balances, to EBITDA, also including IFRS 16" (H1-26 report, p.31). We did not invent one.
- **June ratios are not comparable with the year-end cap.** Dividends (€1.4bn) and working capital (−€1.2bn) both hit in H1. That is why the model rolls forward from June to December and does not apply 3x to the June debt figure.
- **Maturities.** About €3.9bn of contractual flows fall in 2027 (FY25 statements, p.66). In H1-26 Veolia had already issued €3.9bn of bonds. Liquidity is €15.2bn gross and €3.8bn net of current debt. **Refinancing does not constrain capacity. Leverage does.**
- **No financial covenant at parent level** (FY25 statements, p.67). The 3x is a public commitment, not a contractual one. The hard backstop is the rating (BBB / Baa1, both stable, reaffirmed April and May 2026).

---

## 4. The assumptions we must be ready to defend

Yellow cells in the workbook are judgement calls. These are the ones that move the answer. They are ranked by the Tornado sheet, which the model computes.

| Rank | Driver | Tested range | Capacity range | Our base and why |
|---|---|---|---|---|
| 1 | Leverage cap | 2.9x – 3.1x | 568 – 2,138 | **3.0x**: the published commitment. "Slightly above" applies to 2026 only. |
| 2 | EBITDA 2027 level | 7.75 – 8.25bn | 788 – 2,288 | Trajectory **7.94bn before disposals**: FY25 × (1 + 5.5% organic − 0.9% FX) × 1.05, plus Clean Earth. |
| 3 | Disposals cashed by end-2027 | 50% – 100% of €2bn | 943 – 1,763 | **75%**. The programme runs "by mid-2028", and only €73m was cashed in H1-26. |
| 4 | FX / other on net debt 2027 | +€400m / −€400m | 953 – 1,753 | **0**. H1-26 was −€260m, mostly USD. The USD exposure is larger since Clean Earth. |
| 5 | Acquired EBITDA counted in the ratio | 0 – 100% | 1,353 – 1,933 | **0** (conservative). We need Veolia to say whether its ratio is pro forma. |
| 6 | Organic growth 2027 | 4% – 6% | 1,127 – 1,579 | **5%**: GreenUp "around 5%". |
| 7 | Net FCF growth | −3% – +5% | 1,194 – 1,479 | **+1.5%**: the 2023-25 CAGR. Net FCF has been flat at €1.14–1.18bn while EBITDA grew. |
| 8 | Clean Earth EBITDA | multiple 11.5x – 8.5x | 1,230 – 1,480 | **€162m**, derived as (US$3.04bn ÷ 9.8x − US$120m synergies) ÷ FX. |

**The briefing says "you already know" which input dominates. We now prove it.** After the cap itself, the **EBITDA level is the dominant driver**. Each €100m of 2027 EBITDA is worth €300m of capacity. Disposals and FX come next, at about half that weight. No capital-allocation decision by the other groups moves the answer as much as these inputs do.

**Three points a jury will press on:**
1. **"Your EBITDA is below €8bn. Do you not believe the plan?"** Our trajectory gives €7.94bn before disposals. Whether the €8bn target is met depends on Clean Earth's contribution. We also run the €8bn case: it adds about €0.2bn of capacity. We deduct the EBITDA lost through disposals in both cases. If the €8bn target is meant "before disposals", we say so, and the €8.0bn case is the right one.
2. **"Disposals give you cash but take EBITDA away."** Correct, and the model includes this. At 10x, every €1bn sold removes €100m of EBITDA and therefore €300m of debt room. So €1bn sold gives €0.7bn of net capacity. Below about 3x a sale would add nothing, which no plausible deal reaches.
3. **"Why ignore the EBITDA you would buy?"** Because we do not know whether Veolia's year-end ratio is pro forma. If it is, capacity rises by a factor of 1 ÷ (1 − 3/multiple), i.e. ×1.43 at 10x. That is the upside case.

---

## 5. Reconciliation issues found (worth putting in the report)

- **Net debt at 31/12/2025 has two values.** It is €19,657m as the opening of the H1-26 bridge (p.26), and €19,821m in the FY25 statements (p.66) and in the H1-26 key figures table (p.6). The €164m gap looks like the Suez PPA remeasurement of financial liabilities. That is consistent with the H1-26 debt table, which shows −143 for this item at Jun-26. **We need to confirm this.**
- **Parent dividend paid in 2025.** It is €1,023m in the H1-25 column but €895m in the FY25 column of the same key figures table (H1-26 report, p.6). €895m is the 2024 amount, so this looks like a carry-over error.
- **"Dividends paid" in Veolia's net debt bridge includes minority dividends and hybrid coupons.** H1-26: 1,307 + 87 = 1,394. H1-25: 1,173 + 94 = 1,267. These agree with the cash flow statements. Our outflow line follows the same definition.
- **The €850m hybrid redeemed in February 2026** was reclassified as debt at 31/12/2025 (FY25 p.74), so it does not add to net debt again in 2026. Rating agencies give hybrids about 50% equity credit. A further redemption without replacement would therefore tighten the rating metrics even with Veolia's net debt unchanged.
- **Hazardous waste target.** It is 9 Mt in the briefing (p.20) and 10 Mt in the GreenUp launch text (H1-24 p.12). This belongs to the other perimeters, but we should flag it to them.

---

## 6. Block 3: the rating trigger, the limit we have not yet nailed down

According to search-result extracts of Moody's credit opinions published on veolia.com, Moody's guidance for Baa1 is FFO/net debt around 20%. The downgrade trigger is "below the high teens". The forecast for 2027 is about 19.5%. Treating "high teens" as 18%, Veolia could take on about 8% more debt before the trigger, roughly €1.8bn. That is slightly more than the €1.35bn of 3x headroom. **So on this first reading, the 3x commitment binds before the rating does.** Confidence is low. This environment blocked downloading the PDFs, and Moody's adjusts debt (pensions, hybrids, leases) in ways we cannot rebuild from Veolia's figures. **Action:** download and attach these documents:
- [Moody's Credit Opinion, 4 May 2026](https://www.veolia.com/sites/g/files/dvc4206/files/document/2026/05/Credit_Opinion-Veolia-Environnement-SA-04May2026.pdf)
- [Moody's Credit Opinion, 2 Dec 2025](https://www.veolia.com/sites/g/files/dvc4206/files/document/2025/12/Credit_Opinion_Veolia-Environnement-2Dec2025.pdf)
- the S&P research update of 27 April 2026

Then replace the two `Moody_*` inputs.

---

## 7. What we could not establish, and the questions for Veolia on 16 October

Evidence gaps (to attach before 17 November):
- FY2025 and FY2024 URD chapter 3: full-year net debt bridges and net financial investments on Veolia's definition, plus the published year-end leverage ratios.
- The FY2023 URD, if the envelope starts in 2023.
- Enviri's 10-K for Clean Earth's real EBITDA and capex, to replace the figure we derived from the multiple. The Subject 1 groups are building this baseline, so we can share it.
- The rating agency reports (section 6).

Questions for the partner session:
1. Is the €4bn GreenUp envelope gross or net of disposals? Does it cover 2023 or start in 2024? Does it include discretionary capex?
2. Is the year-end leverage ratio computed on reported or pro forma EBITDA for acquisitions made during the year?
3. Is the €2bn+ disposal programme counted from November 2025? Do the H1-26 disposals (€73m) count towards it? Roughly what EBITDA multiple are the assets expected to fetch?
4. Is the €8bn EBITDA target for 2027 before or after the disposal programme and Clean Earth?
5. Which metric does Veolia manage to for the rating: 3x net debt / EBITDA, or the agencies' FFO/net debt? What internal buffer does it keep below the trigger?
6. What explains the €164m gap between the two 31/12/2025 net debt figures?

---

## 8. What would change our mind

- **Capacity would rise above €2bn** if Veolia confirms a pro forma ratio, or if disposals are fully cashed by end-2027.
- **Capacity would fall below €0.5bn** if EBITDA lands at €7.75bn or less, if half the disposals slip into 2028, or if the dollar weakens enough to add €0.4bn to net debt.
- **The rating would become the binding limit** if the agency PDFs show a trigger at 19% or above, or a forecast below 19.5%.

*AI use (for the annexe): Claude extracted figures from the four PDFs, located the page references, and wrote the Python script that builds the formula-driven workbook. Every result above is calculated by the workbook. None was typed in.*
