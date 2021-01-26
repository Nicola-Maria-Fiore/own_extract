rename _all,lower
drop if investorfullname=="NULL"

rename A RIC
rename L investor_turnover_percentage
rename O investor_portfolio_number_of_securities_sold

gen holdingsfilingdate_new=.
replace holdingsfilingdate_new = date(holdingsfilingdate, "MDY")
drop holdingsfilingdate
rename holdingsfilingdate_new holdingsfilingdate
format holdingsfilingdate %td

gen earliestholdingsdate_new=.
replace earliestholdingsdate_new = date(earliestholdingsdate, "MDY")
drop earliestholdingsdate
rename earliestholdingsdate_new earliestholdingsdate
format earliestholdingsdate %td

save "ownership.dta"