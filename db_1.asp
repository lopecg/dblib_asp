<html>
<head>
</head>

<body>

<%

accessdb=server.mappath("sampltrc.mdb")
mySQL="select * from [Company]"

strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
strconn=strconn & accessDB & ";"


call query2table(mySQL,strconn)

strconn2="Driver={Microsoft Access Driver (*.mdb)};" & _ 
"Dbq=WestIWMD.mdb;" & _ 
"Uid=admin;" & _ 
"Pwd=" 

call query2count(strconn2)

%>
<!--#include file="lib_dbtable.asp"-->


</body>
</html>
