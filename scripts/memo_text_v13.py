# -*- coding: utf-8 -*-
"""v13 memo text. Numbers tie to NCLH_model_v6.xlsx (Thesis to KPIs, Model vs Consensus, DCF)."""

# (label, [paragraphs], chart key or None)
PAGE1 = [
("Narrative Summary:", [
"Recommendation: short NCLH at $15.05, starting at a third of a full position. Our base case is worth about $10.80 (−28%) and the "
"probability-weighted value about $12 (−21%), against +32% if consensus is right. The market prices NCLH on net yield (ticket and onboard "
"revenue, less commissions and onboard costs, per available berth-day), and the share price assumes net yield climbs back above its 2025 "
"level. We think it falls again in 2027, for two reasons we can measure today. First, NCL is cutting fares to fill Caribbean ships where it "
"is losing share to Royal Caribbean and Carnival, and those cuts land on the cabins it still has to sell for 1H27. Second, NCL’s Best Price "
"Guarantee gives guests little reason to book early, so the booking shortfall that deposits show today is likely to last into 2H27, filled "
"late at a discount or left empty. In our model the two theses take about $140M off 2027 EBITDA through two KPIs, ticket price per passenger "
"day and occupancy. With onboard spending and unit costs at recent run-rates, we get 2027 EBITDA of $2.47B against a Street-implied $2.69B, "
"and EPS of $0.95 against $1.46–1.67."
], None),

("Company Overview:", [
"NCLH runs 35 ships with about 75,000 berths: Norwegian Cruise Line (NCL), the mass-market brand with about 83% of berths, plus the "
"upper-premium Oceania and luxury Regent brands. In 2025 it earned $9.8B of revenue, $2.73B of adjusted EBITDA and $2.11 of EPS, and it "
"carries $14.8B of net debt, two-thirds of enterprise value. The debt is why small changes in net yield move the equity: each 1% is about "
"$75M of EBITDA and 9% of the share price. 2026 went wrong quickly. NCL added about 40% more Caribbean capacity before its Great Stirrup Cay "
"upgrade was ready, the Iran war pushed oil above $100, and guidance fell from flat net yield and $2.38 of EPS in March to about −5% and "
"$1.50 in July. In February John Chidsey replaced Harry Sommer as CEO, under a cooperation agreement with Elliott, a holder of more than 10%."
], None),

("Industry Overview:", [
"Carnival (about 32% of capacity), Royal Caribbean (26%), MSC and NCLH (about 10% each) carry about 80% of the market. Demand is healthy "
"overall: Carnival’s deposits are up 6.5% on flat capacity and Royal’s are up 5.6%. The competition is in the Caribbean, where each large "
"line now sells a private island, and NCL arrived with the most new capacity and the last island to be finished (September 2026). Returns "
"show who is winning: 2025 return on capital was about 16% at Royal and 12% at Carnival, against about 10% at NCLH, close to its cost of capital."
], None),

("Market Expectations:", [
"Since 2024 NCLH’s EV/EBITDA has tracked net yield growth closely (r = 0.87 over ten quarters), and on the four days management changed its "
"yield outlook the stock lost 27 points against Royal and Carnival. Working back from the $21.7B enterprise value at a 9% WACC, the price "
"needs about $3.0B of steady-state EBITDA, which is management’s March guide and net yield back above its 2025 level (Exhibit 1). The Street "
"is less demanding but still expects a recovery: consensus 2027 EPS of $1.46–1.67 implies about $2.7B of EBITDA, and the $18.80 average "
"target needs $2.9B.",
"Sentiment has turned faster than estimates. Ratings are 4 Buy, 15 Hold and 0 Sell (11, 4 and 0 a year ago) and short interest is 20.7% of "
"float, yet consensus 2027 EPS moved only two cents when NCLH guided 2027 interest $0.35 a share higher on September 30. The bull argument is "
"that the selloff was oil and the sector. Our regression puts 54% of the decline there, but NCLH moved no more than Carnival on the biggest oil "
"days, and the NCL-specific 46% came on six company news days and did not reverse."
], "ex1"),

("Thesis 1: NCL is cutting prices because it is losing share in the Caribbean, and the cuts reach 1H27 revenue.", [
"The Street reads the discounting as a result of the war and of execution problems that the finished island and new ships will fix in 2027. "
"The fare data points to competition instead. On October 2 we compared the lowest fare on 4,783 itineraries with each one’s 90-day average. "
"NCL has cut 1H27 Caribbean sailings by a median 14.5% but Europe by only 2%, although two-thirds of its European guests fly in and face "
"higher airfares. Royal and Carnival have held Caribbean prices and are discounting Europe instead (Exhibit 2). If the war were the cause, we "
"would expect the opposite pattern. The cuts are also recent and the new product is not stopping them: NCL went from 1 to 9 of the industry’s "
"10 largest monthly price cuts between December and October, sailings that call at its island are cut as much as its other Caribbean "
"sailings, and its newest ships are cut 22% against 1% for Royal’s newest. Oceania and Regent show no cuts, so the problem is the NCL brand.",
"For the model, the cuts matter only for cabins NCL has not yet sold. With a normal booking curve and the booking shortfall from Thesis 2, "
"about 45% of 1Q27 and 55% of 2Q27 cabins are still unsold, and 74% of NCL-brand capacity sails in the Caribbean in 1Q27. Weighting the NCL "
"brand at 65% of ticket revenue, 1H27 ticket revenue per passenger day (KPI 1) falls 2.6% y/y in our model against 1.1% in the consensus case. "
"The risk to this view is that the cuts reverse before guests book. We re-run the fare panel monthly; cuts back near peers’ levels would undo it."
], "ex2"),
]

PAGE2 = [
("Thesis 2: The price guarantee keeps guests from booking early, so 2H27 does not recover the way consensus expects.", [
"The consensus case has net yield up about 3% in 2H27, helped by easy comparisons and by base-loading: opening sailings at lower fares and "
"raising them as ships fill. That only works if guests expect prices to rise. NCL’s Best Price Guarantee lets a booked guest move to any lower "
"fare before final payment and gives a paid-up guest an upgrade or future-cruise credit afterwards; Royal offers nothing after final payment. "
"With NCL cutting 77% of its 4Q26 Caribbean sailings of seven nights or more by 20% or more, guests have learned that waiting costs them nothing.",
"Deposits are where this shows up first. Deposits per future berth-day are down 8.3% y/y, the seventh decline in a row, while Royal’s and "
"Carnival’s are growing (Exhibit 3). If booked fares fell as much as realized fares (−5.6%), the shortfall in cabins booked is about 3 points; "
"if booked fares were flat, about 8. We use 5. We assume 80% of those cabins sell late at NCL’s current close-in discount (about 25%) and 20% "
"sail empty, about half of what the 2026 link between deposits and occupancy implies. That puts 2027 occupancy (KPI 2) at 102.1%, down again "
"from 2026 and a point below the consensus case, and 2H27 ticket price at −0.6% against +0.3%. The test is the November 4 10-Q: September 30 "
"advance ticket sales below $2.97B support this, and at $3.11B or more we would cover."
], "ex3"),
]

SPLIT_LABEL = "From the data to the model:"
SPLIT_TEXT = (
"Net yield is occupancy times revenue per passenger day, so the theses enter the model in two places: Thesis 1 moves 1H27 ticket price, and "
"Thesis 2 moves 2H27 ticket price and 2027 occupancy. Together they take $137M off 2027 EBITDA (Exhibit 4). The rest of our gap to the "
"consensus case is not part of either thesis. Onboard spending grows 2% in our model (it fell 0.3% in 2025 and rose about 1% in 1H26) against "
"3.5%, and unit costs rise 1% against a 0.5% decline. On the two theses alone, the shares are worth about $13.50 (−10%).")
KPI_TABLE = [
    ["Driver (2027)", "Consensus case", "Ours", "EBITDA ($M)"],
    ["KPI 1: 1H27 ticket price y/y (Thesis 1)", "−1.1%", "−2.6%", "−44"],
    ["KPI 1: 2H27 ticket price y/y (Thesis 2)", "+0.3%", "−0.6%", "−17"],
    ["KPI 2: occupancy (Thesis 2)", "103.1%", "102.1%", "−76"],
    ["Other: onboard spend per guest y/y", "+3.5%", "+2.0%", "−42"],
    ["Other: unit costs y/y / fuel per ton", "−0.5% / $760", "+1.0% / $800", "−107"],
    ["2H26 exit rate", "", "", "−29"],
    ["2027 EBITDA", "$2.79B", "$2.47B", "−315"],
]

AFTER = [
("Valuation and risk/reward:", [
"We value NCLH on 2027 EBITDA at today’s 8.4x forward multiple, less $15.8B of year-end net debt. Our $2.47B gives $10.83 a share (−28%), "
"and a DCF of 2027–30 cash flows gives $8.11. Leverage widens the range of outcomes: each point of 2027 ticket price is worth about $0.90 a "
"share and each point of occupancy about $1.40. If consensus is right and the stock re-rates to 9x, it is worth about $19.90 (+32%); if "
"discounting spreads into 2H27, about $6.30 (−58%). Weighted 25/50/25, value is about $12 (−21%). Net debt also keeps rising in our model, "
"to $17.7B (7.2x EBITDA) at the end of 2027, as newbuild payments continue."
], "table"),
("Catalysts:", [
"November 4 (3Q26 results): earnings should beat ($0.90 guided against $0.83 consensus), so the number to watch is September 30 advance ticket "
"sales, which we expect at $2.94B (−7% y/y). Late February: the first 2027 guide, which we expect at or below flat net yield against +1.6% in "
"the consensus case. May 2027 (1Q27 results): we model occupancy of 101.8% and ticket price per passenger day down 2.8%."
], None),
("Risks and sizing:", [
"The main risk is a squeeze. 95M shares are short (5.7 days to cover), and the stock rose 48% in six weeks this spring when oil fell. About "
"half of NCLH’s moves follow oil and its peers, so we would start with a third of the position, add on oil- or ceasefire-driven rallies and "
"around November 4 and February, hold part of it in puts, and stop out above $19, where shorts added at about $18 are under water and the "
"bull case begins. Other risks: Elliott’s standstill ends in February, cost savings beyond the $225M identified would offset our cost "
"assumption, and the fare panel is a snapshot, so the cuts could reverse."
], None),
]

SCEN_TABLE = [
    ["Scenario", "Prob.", "2027 net yield", "2027 EBITDA", "EV / EBITDA", "Value", "vs. $15.05"],
    ["Bull: consensus is right and the stock re-rates", "25%", "+1.6%", "$2.79B", "9.0x", "$19.92", "+32%"],
    ["Base: our view", "50%", "−0.7%", "$2.47B", "8.4x", "$10.83", "−28%"],
    ["Bear: discounting spreads into 2H27", "25%", "−2.6%", "$2.22B", "8.4x", "$6.27", "−58%"],
    ["Probability-weighted", "", "", "", "", "$11.96", "−21%"],
]

SOURCES = ("Sources: NCLH 8-Ks (Nov 4, 2025–Sep 30, 2026), 10-Ks FY24–25, 2Q26 10-Q and call; ncl.com Best Price Guarantee terms; "
"allaboarddeals.com fares (collected Oct 2, 2026) and policy summary (Apr 2026); Carnival 8-K (Sep 29, 2026); Royal Caribbean 2Q26 10-Q; SEC XBRL; "
"Nasdaq/Zacks and Yahoo consensus (Oct 2, 2026); Wells Fargo action via news reports (Sep 2026); FINRA short interest; FRED; Kenneth French Data "
"Library. Estimates are the team’s (NCLH_model_v6.xlsx). AI disclosure: [team to complete under Rule 10].")
