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

accessdb2=server.mappath("WestIWMD.mdb")

strconn2="Driver={Microsoft Access Driver (*.mdb)};" & _ 
"Dbq=" & accessdb2 & ";" & _ 
"Uid=admin;" & _ 
"Pwd=" 

mySQL2="Select * FROM Personnel"
response.Write("<br>")
call query2count(mySQL2, strconn2)

Dim catDB ' As ADOX.Catalog
		Set catDB = Server.CreateObject("ADOX.Catalog")
		' Open the catalog
		catDB.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
		"Data Source=" & accessdb2
%>
<!--#include file="libdbtable.asp"-->

call createtable("Contactos", catDB)

</body>
</html>
