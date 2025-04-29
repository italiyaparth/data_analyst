data range name = VData
array criteria range name = LotNo
additional single criteria = B2
additional single criteria = C1


=LET(data,FILTER(VData,( CHOOSECOLS(VData,8)=B2 )*( CHOOSECOLS(VData,9)=C1 )*ISNUMBER( MATCH(CHOOSECOLS(VData,1),LotNo,0) ) ),IFERROR(SUM(CHOOSECOLS(data,7)),""))