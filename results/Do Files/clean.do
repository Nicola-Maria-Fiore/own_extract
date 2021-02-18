*rename vars
rename a ric
rename l investor_turnover_percentage
rename o inv_port_n_secs_sold
*rename other vars?

*label vars
label var ric "RIC"
label var investor_turnover_percentage "Investor Turnover Percentage"
label var inv_port_n_secs_sold "Investor Portfolio Number of Securities Sold"

*replace "" with "."
ds, has(type string)
foreach v in `r(varlist)' {
	replace `v' ="." if `v'=="NULL"
	}

*destring all vars
destring, replace

*generate date vars
gen holdings_filing_date_new=.
replace holdings_filing_date = date(holdingsfilingdate, "MDY")
format holdings_filing_date %td
gen earliest_holdings_date_new=.
replace earliest_holdings_date = date(earliestholdingsdate, "MDY")
format earliest_holdings_date %td

save "ownership_clean.dta"


