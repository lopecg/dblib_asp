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
call query2count(strconn)

%>
<!--#include file="lib_dbtable.asp"-->


</body>
</html>
