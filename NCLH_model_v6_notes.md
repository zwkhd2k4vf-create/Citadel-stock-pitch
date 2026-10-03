# NCLH model v6 and memo v13: what changed and why

This covers what changed from model v5 / memo v12 to model v6 / memo v13. It's organized around the judging criteria: whether the thesis holds up, whether the model and thesis line up, and whether the model's assumptions make sense.

## The chain the model now shows

The **Thesis to KPIs** tab (now the second tab) shows each thesis as a chain: alternative data → what it implies → the KPI it moves → consensus case vs. ours → 2027 EBITDA.

Net yield = occupancy × revenue per passenger day, so the theses enter the model in two places:

| | Data | KPI | Consensus case | Ours | 2027 EBITDA |
|---|---|---|---|---|---|
| Thesis 1 | Fare panel: NCL 1H27 Caribbean −14.5% vs. peers −0.5% | KPI 1: 1H27 ticket price y/y | −1.1% | −2.6% | −$44M |
| Thesis 2 | Deposits per future berth-day −8.3% y/y | KPI 1: 2H27 ticket price y/y | +0.3% | −0.6% | −$17M |
| Thesis 2 | Same (booking shortfall) | KPI 2: 2027 occupancy | 103.1% | 102.1% | −$76M |
| Other | Onboard trend (2025 −0.3%, 1H26 +1%) | Onboard spend y/y | +3.5% | +2.0% | −$42M |
| Other | Run-rate costs | Unit costs / fuel | −0.5% / $760 | +1.0% / $800 | −$107M |
| | | 2H26 exit | | | −$29M |

The bridge is now live: the **Bridge Build** tab reruns the quarterly net yield one input group at a time. Its checks against `DCF!E41` and `DCF!E42` are both 0.

## New and restructured tabs

- **KPI 1 Price Build (Thesis 1):** fare cuts applied only to cabins still unsold, by brand (NCL vs. Oceania/Regent) and region (Caribbean vs. other).
  - The observed cuts are now formulas over the raw fare panel. In v5 they were typed in.
  - Includes a sensitivity table.
- **KPI 2 Booking Build (Thesis 2):** the booking shortfall from deposits, the share of it that sails empty (occupancy), and the late discount on the rest (2H27 price). The price-guarantee claim sensitivity moved here.
- **Fleet Qtr Build:** ship × quarter capacity days, 504 rows. 2027 quarterly capacity growth now comes from delivery and exit dates instead of typed-in numbers.
- **Fare Panel:** the old *Alt Data Build*, now holding data only. Section H holds the summary statistics the memo quotes. The duplicate *Alt Data* tab is gone.

## Deposit lag (KPI 2 Booking Build, section F)

Since COVID, deposits lead net yield by about four quarters, not the one to two quarters seen before.
- **Why:** guests book further ahead. Advance ticket sales were about 25% of annual revenue in 2016–18 and about 34% in 2023–25.
- **Track record:** at a four-quarter lead, the indicator called the direction of net yield in 7 of 8 quarters. That is 5 of 6 counting reported quarters only; the other two outcomes are the 3Q26 guide and 4Q26 implied figure. At a one-quarter lead it called 6 of 10.
- **What it implies for 1H27:** a simple fit over those eight quarters gives 1H27 net yield of about −3.9%, against −2.6% in our base case and −0.2% in the consensus case. This is a cross-check only, not a model input. Memo Exhibit 3 plots deposits shifted forward four quarters.

## Logic fixes that moved the numbers

1. **Brand weighting.** v5 applied NCL's fare cuts to all NCLH capacity, including Oceania and Regent, whose fares show no cuts. That overstated Thesis 1 by about a third.
   - v6 weights the NCL brand at 65% of ticket revenue. This is an input: the berth share of 83% is the ceiling, and the fare panel implies 58%.
2. **The booking shortfall now comes from deposits.**
   - Deposits per future berth-day are −8.3% y/y. If booked fares fell like realized fares (−5.6%), about 3 points of the drop is a shortfall in cabins booked; if booked fares were flat, about 8. We use 5 points.
   - v5 used 10 points for 1Q27 with no derivation.
3. **Occupancy and the late discount come from one rule.** 20% of the shortfall sails empty and 80% sells late at NCL's observed close-in cut (−25%). The empty share is about half what the 2026 deposit–occupancy relationship implies.
   - These replace v5's typed-in occupancy changes and its "base-loading" plug.
4. **One unexplained difference from consensus is gone.** v5 had 2027 underlying pricing 0.5 pts below the consensus case with no reason given; v6 sets it equal.
5. **2026 is unchanged.** The 4Q26 inputs were recalibrated, so the 2026 net yield path (−4.7%, 4Q26 −7.0%) and 2026 EBITDA ($2.47B) still match v5 and guidance.
   - The consensus case was left at a 25% exposure to today's cuts and not re-tuned. Its 2027 EPS of about $1.61 is still inside the $1.46–1.67 consensus range.

## Results: v5 → v6

| | v5 | v6 |
|---|---|---|
| 2027 EBITDA | $2.44B | **$2.47B** |
| 2027 EPS | $0.88 | **$0.95** |
| 2027 net yield | −1.3% | **−0.7%** |
| Base value | $10.22 (−32%) | **$10.83 (−28%)** |
| Probability-weighted | $11.67 (−22%) | **$11.96 (−21%)** |
| Bull / bear | $19.78 / $6.46 | **$19.92 / $6.27** |
| Theses-only value (consensus onboard, costs, fuel) | n/a | **$13.52 (−10%)** |

## Decision for the team

With honest logic, the two theses explain **$137M of the $315M** gap to the consensus case. The rest is costs (+1% vs. −0.5%), onboard spend and the 2H26 exit.

The memo now says this openly and frames the trade as risk/reward with staged sizing. If you want the theses to carry more of the gap, the defensible levers are:

- **Shortfall:** use 8 points (the fares-flat reading) instead of 5.
- **Share that sails empty:** use 30%+ (the full 2026 relationship) instead of 20%.

Both are single input cells on the KPI 2 tab. Change them only if you can defend the change in Q&A.

## Still open from the v5 review

- EV/EBITDA uses a hard-coded $15.8B year-end net debt, while the DCF uses the model's $15.88B (about $0.18/share).
- **Raw itinerary file.** The fare panel in the workbook is a summary by line, region and window. A true itinerary-level tab (4,783 rows) needs the raw scrape; send the CSV and it can be rebuilt bottom-up.

## Final pass

- **The gap is now all revenue.** The consensus case reproduces Street 2027 revenue ($10.62B vs. $10.57B). At run-rate unit costs it also matches Street EBITDA ($2.68B vs. $2.69B).
  - From there, the full $208M gap to our $2.47B comes from the two data sets:
    - Fare panel: 4Q26 exit −$29M and 1H27 ticket price −$44M.
    - Deposits: 2H27 price −$17M and occupancy −$76M.
    - Guest mix of discount-filled ships: onboard spend −$42M.
  - Consensus EPS also needs unit costs to fall 0.5% (another $107M). The Street's own EBITDA doesn't.
  - Model vs Consensus rows 24–35 and memo Exhibit 4 show this.
- **NCL brand weight is now calibrated (KPI 1 Build section H).** If Oceania and Regent sell near their lowest fares, NCL is 66% of the reported $265 of 2025 ticket revenue per passenger day. We use 65%.
- **Year-end net debt is now linked to the model** ($15.88B, previously typed in as $15.8B). The forward multiple is 8.46x.
  - New results: base $10.82 (−28%), bull $19.74, bear $6.24, probability-weighted $11.90 (−21%), DCF $8.26.
- **Critiques and answers:** Thesis to KPIs section F covers ten likely critiques. They include lowest vs. realized fares, war vs. competition, brand weight, the short-cruise mix, the deposit lag, the gap to consensus, how the shortfall is sized, the multiple (9.0x gives $13.67), cost savings ($12.76 if unit costs fall 0.5%) and the crowded short.
