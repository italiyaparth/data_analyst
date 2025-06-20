=LET(
StartDate,$H$3,
EndDate,$J$3,
SellerName,$D49,

data,FILTER(Sales_Pivot_Customer,(Sales_Pivot_Sales_Date>=StartDate)*(Sales_Pivot_Sales_Date<=EndDate)*(Sales_Pivot_Assigned_New_Seller_Name=SellerName),"Empty"),

IF(TYPE(data)=2,"",COUNTA(UNIQUE(data)))
)




 
- In above Formula, TYPE(data)=2    Means   we have written "Empty" (a String) IF Array is Empty in FILTER    Means    data is NOT an Array (it is an Empty Array)

- If we use "IF(data='Empty','',COUNTA(UNIQUE(data)))"; 
	then IF formula will apply to all rows of data which will give same answer in all rows; 
	For Example, Answer=3 --> row1=2, row2=2, row3=3



- Only for COUNT or SUM or .etc   after  FILTER formula,   this is a important point to remember; 
	otherwise we can use "IF(data=0,"",data)" in direct SUMIFS, COUNTIFS, .etc formulas