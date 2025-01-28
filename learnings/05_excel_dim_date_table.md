1.	Open a blank query in Power Query of Power BI
2.	type ={Number.From(#date(2018,1,1))..Number.From(#date(2018,12,31))}
3.	That will generate a series of numbers as a list
4.	Convert it to a table (right click, select to table option).
5.	Convert the ABC123 type to date
6.	Rename to Date.
7.	Now you have a date table. Add columns as necessary (year, month, month name, etc) to make your date table suit your needs.
8.	To add start of the month, click on date option arrow in add column menu in ribbon
9.	Same as above, add year Column, convert it to text
10.	Suppose, fiscal year 2021 is 1-sep-2020 to 31-aug-2021
11.	So, add new custom column, fy_month=Date.addMonths([month], 4) 
12.	Add new custom column, fy=Date.Year([fy_month]) 
13.	Add new custom column, to display real month as MMM
14.	Add new custom column, to display fiscal month no. 
Now, we want to add quarter
●	Add new custom column, fy_qtr=”q”&Roundup(date.month([fy_month]/3)) 



If Error:

1.	Custom Column
2.	= Date.toText([date], “yyyy”) & “/” & Text.PadStart(Text.From(Date.Day([date])), 2,”0”) & “/” & Text.PadStart(Text.From(Date.Month([date], 2,”0”) 
3.	Convert to date

