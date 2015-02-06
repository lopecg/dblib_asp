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

mySQL2="Select * FROM Personnel"
response.Write("<br>")
call query2count(mySQL2, strconn2)

%>
<!--#include file="libdbtable.asp"-->


</body>
</html>
