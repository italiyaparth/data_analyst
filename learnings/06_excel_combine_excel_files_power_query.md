Source : 


Custom1 : 

	Table.TransformColumns(
		Source, {"Content", each Excel.Workbook(_, true)}
	)


Expanded Content :


Data : 
	Table.Combine(#"Expanded Content"[Data])
