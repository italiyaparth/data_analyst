
shp,$X$80 = RB
col,$Y$80:$Y$82 = D | E | F
clr,$Z$80:$Z$84 = IF | VVS1 | VVS2 | VS1 | VS2
lab,$AA$80 = IGI
colTinge,$AB$80 = blank
bicNOTselect,$AC$80 = B2
bisNOTselect,$AD$80:$AD$81 = B1 | B2



=LET(
     shpColumn,Data!$B$2:$B$322636,
     ctsColumn,Data!$C$2:$C$322636,
     colColumn,Data!$D$2:$D$322636,
     clrColumn,Data!$E$2:$E$322636,
     labColumn,Data!$F$2:$F$322636,
     colTingeColumn,Data!$G$2:$G$322636,
     bicColumn,Data!$H$2:$H$322636,
     bisColumn,Data!$I$2:$I$322636,

     shp,$X$80,
     col,$Y$80:$Y$82,
     clr,$Z$80:$Z$84,
     lab,$AA$80,
     colTinge,$AB$80,
     bicNOTselect,$AC$80,
     bisNOTselect,$AD$80:$AD$81,

     SUMPRODUCT(ctsColumn,
           --( shpColumn = shp ),
           --( ISNUMBER(MATCH(colColumn, col, 0)) ),
           --( ISNUMBER(MATCH(clrColumn, clr, 0)) ),
           --( labColumn = lab ),
           --( colTingeColumn = colTinge ),
           --( bicColumn <> bicNOTselect ),
           --( ISNA(MATCH(bisColumn, bisNOTselect, 0)) )
     )
)




- --( bicColumn <> bicNOTselect) => --( NOT( bicColumn = bicNOTselect) ) 


- use MAP for P&L => MAP(sales, cost, LAMBDA(a, b, a/b-1)) 