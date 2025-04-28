Source : 


Custom1 : 

	Table.TransformColumns(
		Source, {"Content", each Excel.Workbook(_, true)}
	)


Expanded Content :

here, insert a step to filter only visible files (removing HIDDEN files)
(Do NOT keep open these linked files when refreshing main file, because "~$...." file enters below as well)

Data : 
	Table.Combine(#"Expanded Content"[Data])
