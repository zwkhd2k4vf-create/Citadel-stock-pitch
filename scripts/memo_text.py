# -*- coding: utf-8 -*-
"""v12 memo text. Every number ties to NCLH_model_v5.xlsx (see model_check.md)."""

# (label, [paragraphs], chart rId or None)
SECTIONS = [
("Narrative Summary:", [
"NCLH trades on net yield, which is ticket and onboard revenue, less commissions and onboard costs, per available berth-day. "
"Each 1% of net yield is worth about $75M of EBITDA and, with net debt at two-thirds of enterprise value, about 9% of the equity. "
"Net yield rose 2.3% in 2025 and is guided down about 5% for 2026. Only about a quarter of that drop is lower occupancy; the rest is lower prices per guest. "
"The share price assumes net yield gets back above its 2025 level, and a case calibrated to consensus has it up 1.4% in 2027. "
"We expect it to fall another 1.3%. First, NCL is cutting prices because it is losing share to Royal Caribbean and Carnival in the Caribbean, "
"so the cuts will not end with the war. Its new private island and new ships are being discounted as heavily as the rest of the fleet. "
"Second, management plans to rebuild prices in 2H27 by opening sailings at lower fares and raising them as ships fill. "
"That depends on guests booking early, and NCL’s Best Price Guarantee gives them little reason to, because guests who have already booked are compensated when NCL cuts a fare. "
"We forecast 2027 EBITDA of $2.44B against a Street-implied $2.69B, and EPS of $0.88 against consensus of $1.46–1.67. "
"At today’s multiple that is worth $10.22 a share, 32% below the current $15.05."
], None),

("Company Overview:", [
"NCLH runs 35 ships with about 75,000 berths. Norwegian Cruise Line (NCL), the mass-market brand, has most of the capacity; "
"Oceania (upper premium) and Regent Seven Seas (luxury) have the rest. In 2025 NCLH had $9.8B of revenue (68% tickets, 32% onboard), "
"$2.73B of adjusted EBITDA and $2.11 of adjusted EPS. Net debt was $14.8B at June 30, or 5.4x 2025 EBITDA. "
"Three things hurt 2026. NCL added about 40% more Caribbean capacity for early 2026, before its Great Stirrup Cay upgrade was ready; "
"it began paying travel advisors commission on the full fare for bookings made from December 26, 2025; and the Iran war in February pushed oil above $100. "
"Guidance went from flat net yield and $2.38 of EPS in March to about −5% and ~$1.50 in July. "
"Net yield was −1.0% in 1Q26 and −2.6% in 2Q26, 3Q26 is guided to −8.9% (NCLH said on September 30 it would beat that), "
"and management’s comments imply about −6.5% for 4Q26. In February John Chidsey, a former CEO of Subway and Burger King, replaced Harry Sommer as CEO. "
"He is also chairman, under a cooperation agreement with Elliott, which owns more than 10%."
], None),

("Industry Overview:", [
"Carnival (about 32% of global capacity), Royal Caribbean (26%), MSC (10%) and NCLH (10%) carry about 80% of the market. "
"Demand is at record levels. Carnival reports record booked occupancy and pricing for 2027 and deposits up 6.5% on flat capacity, and Royal’s deposits are up 5.6%. "
"Returns differ much more. In 2025 return on invested capital was about 16% at Royal and 12% at Carnival, against about 10% at NCLH, close to its cost of capital. "
"Most of the competition is in the Caribbean, where the big lines now compete on private islands: Royal’s CocoCay and its Nassau beach club (opened December 2025), "
"and Carnival’s Celebration Key (opened July 2025, about 2.5M guests in its first year). "
"NCL brought the most new Caribbean capacity and was the last to finish its island (September 2026). "
"A ship costs about the same to sail full or empty, so when every line adds capacity at once, the one with the weakest product fills its ships by cutting price."
], None),

("Market Expectations:", [
"NCLH’s valuation has tracked net yield closely since 2024. Over the last ten quarters its EV/EBITDA has had a 0.87 correlation with net yield growth. "
"On the four days management changed its yield outlook the stock lost 27 points against Royal and Carnival, more than its entire company-specific decline. "
"Working back from the $21.7B enterprise value at a 9% WACC, the price needs about $3.0B of steady-state EBITDA. "
"That is management’s March 2026 guide before two cuts, and it would take net yield back above its 2025 level. "
"Consensus is lower but still assumes a recovery. Its 2027 EPS of $1.46 (Zacks) to $1.67 (Yahoo) implies about $2.7B of EBITDA, "
"and the $18.80 average target (+25%) needs $2.9B at today’s 8.4x forward EV/EBITDA. "
"At 8.1x trailing EBITDA the stock looks cheap against 10.9x in 2016–19 and 9.7x since 2024, but that is on EBITDA we expect to fall.",
"Sentiment has turned faster than estimates. Ratings are 4 Buy, 15 Hold and 0 Sell (11, 4 and 0 a year ago), and short interest is 20.7% of float "
"(95M shares, up from 35M in February). Yet when NCLH guided 2027 interest $0.35 a share above 2026 on September 30, 2027 consensus EPS fell only two cents. "
"Analysts are anchored on a company that has met or beaten its quarterly yield guidance twelve times in a row, even while cutting the full year twice. "
"Bulls see the selloff as an oil and sector move. Our regression puts 54% of the decline on the market, oil and the other cruise lines, "
"but NCLH moved no more than Carnival on the biggest oil days, and the NCL-specific 46% came on six company news days and did not reverse. "
"When Brent fell 38% in May and June, NCLH rose 48% but made up only about half of its underperformance against peers, and the July guidance cut erased that."
], "rId6"),

("Thesis 1: NCL is cutting prices because it is losing share in the Caribbean.", [
"The Street sees the discounting as a result of the war and of execution mistakes, and expects the finished island and new ships to fix it in 2027 "
"(Wells Fargo stays Overweight, citing the new Great Tides waterpark). On October 2 we took the lowest fare for 4,783 itineraries from a public fare tracker "
"and compared each with its own 90-day average. NCL has cut first-half 2027 Caribbean sailings by a median 14.5% (45% of them by 20% or more) but Europe by only 2%, "
"although two-thirds of its European guests fly in and US airfares are up 23%. "
"Royal and Carnival have held Caribbean prices (−1% and 0%) and are discounting Europe instead (Celebrity −8.5%).",
"NCL is now cheaper than the budget brand. From Port Canaveral in 1H27 its lowest fare is $94 a night, against $100 at Carnival and $118 at Royal, "
"and 71% of its Caribbean itineraries are at record-low prices, against 45% at both rivals. The change is recent. "
"In 2025 NCL cut prices less often than Royal (cuts of 5% or more on 6% of sailings 101–200 days out, against 24% for Royal). It now has 9 of the industry’s 10 largest monthly price cuts, "
"up from 1 in December, even after the new CEO started and the island was finished. The island and new ships have not helped pricing. "
"Sailings that stop at NCL’s island are cut 13%, the same as its other Caribbean sailings, while Royal’s island sailings are cut 0.5%. "
"NCL’s newest ships, Aqua and Luna, are cut 22% against 1% for Royal’s newest, and its older ships on the same routes are cut 39–42% "
"as Prima-class ships grow from 11% of NCL berths in 2024 to 25% in 2027. "
"Oceania and Regent show no cuts on the same method, so the problem is specific to NCL.",
"In the model we apply the observed cuts only to cabins that are still unsold. For 1Q27 that is about half, based on NCL’s usual booking curve "
"(about 60% sold four to five months out) less the shortfall management has disclosed, and 62% of 1Q27 capacity sails in the Caribbean. "
"That gives net yield of −4.8% in 1Q27 and −3.5% in 1H27, against −0.5% in the consensus case, or $115M of 2027 EBITDA."
], "rId8"),

("Thesis 2: NCL’s price guarantee will keep prices from recovering in 2H27.", [
"The consensus case has net yield up 3.2% in 2H27, helped by easy comparisons and by base-loading, which means opening sailings at lower fares and raising them as ships fill. "
"Chidsey has said NCL was “holding price too high, too far out,” which “left us more exposed to close-in discounting.” "
"Base-loading needs customers to believe prices will rise, so that booking early pays. NCL’s Best Price Guarantee removes that reason. "
"Before final payment a guest can move to any lower fare on the same sailing, and after final payment a guest on a cruise of six nights or more "
"gets an upgrade or a future-cruise credit for the difference. Rivals offer much less. An early booker loses nothing if the price falls, "
"and NCL keeps cutting late. Of its 4Q26 Caribbean sailings of seven nights or more, 77% have been cut 20% or more in the last 90 days (median −34%), "
"after most guests had paid in full, against 7% at Royal and 0% at Carnival.",
"Deposits show the effect first, because guests pay them months before they sail. Deposits per future berth-day are down 8.3% y/y, the seventh decline in a row, "
"and NCL’s first-half deposit build fell 32% to $482M on about 4% more capacity, while Royal’s and Carnival’s deposits grew. "
"Over 20 non-COVID quarters this measure called the direction of net yield over the next two quarters 80% of the time, though this time it turned negative about a year early. "
"Management says NCL is “below its optimal booked position for the next 12 months.” "
"In 1Q26 NCL filled more cabins by cutting price: occupancy rose 2.3 points and net yield still fell 1.0%.",
"In the model, base-loaded fares open 4% lower on a quarter of 2H27 inventory, occupancy recovers half a point less than in the consensus case "
"and onboard spending grows 2% rather than 3.5%. That gives 2H27 net yield of +0.9% against +3.2%, or $89M of 2027 EBITDA, "
"and 2027 net yield of −1.3% against +1.4%. We leave the direct cost of the guarantee out of our base case. "
"If one in three eligible guests claims, it takes another 0.7 points off 2027 net yield, about $54M of EBITDA."
], "rId10"),

("Valuation:", [
"We build net yield by quarter from occupancy, ticket and onboard revenue per guest, and commission and onboard-cost ratios, with capacity built ship by ship. "
"Unit costs excluding fuel rise 1% in 2027 (consensus case: −0.5%), fuel costs $800 a ton (consensus case: $760) and interest is the $870M guided on September 30. "
"That gives 2027 EBITDA of $2.44B, 9.5% below the Street, and EPS of $0.88, below Zacks consensus in every 2027 quarter. "
"Of the $342M gap between the consensus case and ours, $204M comes from the two theses, $30M from a weaker exit from 2026 and $107M from costs and fuel. "
"At today’s 8.4x forward EV/EBITDA, less $15.8B of year-end net debt and divided by 466M diluted shares, that is $10.22 a share (−32%). "
"With the consensus case’s costs and fuel instead of ours, our revenue forecast alone still gives about $12 a share (−19%). "
"Debt is why a 9.5% EBITDA miss turns into a 32% lower share price: each point of 2027 net yield is worth about $1.40 a share. "
"In our model net debt also rises to $17.7B (7.3x EBITDA) by the end of 2027 as newbuild payments continue. "
"A DCF of 2027–30 cash flows at a 9.0% WACC with an 8.4x exit multiple gives $7.31."
], "rId12"),
]

TABLE_AFTER = "Valuation:"

AFTER_TABLE = [
("Catalysts:", [
"NCLH reports 3Q26 on November 4. Earnings should beat (guided $0.90 against $0.83 consensus), so the number to watch is September 30 advance ticket sales in the 10-Q. "
"Last year’s June-to-September pattern points to about $3.0B and the recent trend to about $2.88B; we expect $2.94B, down 7% y/y. "
"A figure below $2.97B supports the thesis, and at $3.11B or above we would cover. In late February NCLH gives its first 2027 guidance. "
"Chidsey’s 2026 bonus is fixed at $2.9M and his 2027 plan resets, so he has reason to guide low, and we expect net yield guidance at or below flat, "
"against +1.4% in the consensus case."
]),
("Risks:", [
"The main risk is a squeeze. Short interest is 95M shares (5.7 days to cover), and the stock rose 48% in six weeks this spring when oil fell. "
"Shorts were added at about $18 on average, so forced covering is most likely in the high teens. "
"About half of NCLH’s moves follow oil and its peers, so we would add on oil- or ceasefire-driven rallies and build the position in thirds around November 4 and February. "
"We would hold part of it in puts and stop out above $19, where the average short is under water and our bull case begins. "
"Other risks are Elliott’s standstill ending in February and cost savings beyond the $225M identified so far, which would offset part of our cost assumption."
]),
]

SOURCES = ("Sources: NCLH 8-Ks (Nov 4, 2025–Sep 30, 2026), 10-Ks FY24–25, 2Q26 10-Q and call; ncl.com Best Price Guarantee terms; "
"allaboarddeals.com fares (collected Oct 2, 2026) and policy summary (Apr 2026); Carnival 8-K (Sep 29, 2026); Royal Caribbean 2Q26 10-Q; SEC XBRL; "
"Nasdaq/Zacks and Yahoo consensus (Oct 2, 2026); Wells Fargo action via news reports (Sep 2026); FINRA short interest; FRED; Kenneth French Data Library. "
"Estimates are the team’s (Excel model). AI disclosure: [team to complete under Rule 10].")

# scenario table: old text -> new text (model DCF!B41:G45)
TABLE_FIX = {"$19.80": "$19.78", "+32%": "+31%", "$10.20": "$10.22", "$6.45": "$6.46", "$11.70": "$11.67"}
