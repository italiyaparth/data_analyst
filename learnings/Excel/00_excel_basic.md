Excel

Code
Char
Clean - for 0to31
Trim - for 32
Substitute - for 127,129,141,143,144,157,160

Index Match
●	When using it in a table, to lock cell reference use Movies[[movie_id]:[movie_id]], 
So there won't be any increment in cell reference when dragging formula to next column
●	example-
=INDEX(Financials,MATCH(Movies[[movie_id]:[movie_id]],Financials[[movie_id]:[movie_id]],0),MATCH(Movies[[#Headers],[budget]],Financials[#Headers],0))
●	In above formula, variables:movie_id, [budget] header

Xlookup
●	Example-
=XLOOKUP(Movies[@[movie_id]:[movie_id]],Financials[[#All],[movie_id]:[movie_id]],Financials[[#All],[budget]],"Not Available",0)

Mean, Median, Mode
●	=AVERAGE
●	=MEDIAN
●	=MODE

Variance, Standard Deviation
●	=VAR.P
●	=STDEV.P
●	Variance is [ SUM of (Xi-MEAN)² ] / total Count
●	Standard Deviation is Square Root of Variance

Correlation
●	=CORREL

EMI monthly
●	=-PMT(rate/12, nper*12, PV) 
●	Rate-annual interest rate
●	Nper-loan period in years
●	Pv-loan amount
