range name = mypointer

B2 = INDEX( mypointer ,MATCH(1,( A2 <=VALUE(RIGHT( mypointer ,LEN( mypointer )-FIND("-", mypointer ))))*( A2 >=VALUE(LEFT( mypointer ,LEN( mypointer )-FIND("-", mypointer )))),0))

B3 = INDEX( mypointer ,MATCH(1,( A3 <=VALUE(RIGHT( mypointer ,LEN( mypointer )-FIND("-", mypointer ))))*( A3 >=VALUE(LEFT( mypointer ,LEN( mypointer )-FIND("-", mypointer )))),0))


--------  OR ---------


B2 = INDEX( mypointer ,MATCH(1,(VALUE(TEXTBEFORE( mypointer ,"-"))<= A2 )*(VALUE(TEXTAFTER( mypointer ,"-"))>= A2 ),0))

B3 = INDEX( mypointer ,MATCH(1,(VALUE(TEXTBEFORE( mypointer ,"-"))<= A3 )*(VALUE(TEXTAFTER( mypointer ,"-"))>= A3 ),0))