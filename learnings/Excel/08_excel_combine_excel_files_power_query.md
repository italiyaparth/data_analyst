Source : 


Custom1 : 

	Table.TransformColumns(
		Source, {"Content", each Excel.Workbook(_, true)}
	)


Expanded Content :

Filtered Rows: here, insert a step to filter only visible files (removing HIDDEN files)
(Do NOT keep open these linked files when refreshing main file, because "~$...." file enters below as well)

if you want to add a custom column which will have filename as a value. it is best when you want to add date as per file. 

Custom2 :

 Table.AddColumn(#"Filtered Rows", "DataWithFileName", each Table.AddColumn([Data], "FileName", (row) => [Name])) 

Data : 
	Table.Combine(#"Expanded Content"[DataWithFileName])


now, to change filename to date, use replace to get 20251231 format then add custom column and use Text.End([FileName],2)&"-"&Text.Middle([FileName], 4,2)&"-"&Text.Start([FileName], 4) 
