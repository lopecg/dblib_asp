<% 
sub query2table(inputquery, inputDSN)
    dim conntemp, rstemp
    set conntemp=server.createobject("adodb.connection")
    conntemp.open inputDSN
    set rstemp=conntemp.execute(inputquery)
    howmanyfields=rstemp.fields.count - 5%>
    <table border=1><tr>
    <% 'Put Headings On The Table of Field Names
    for i=0 to howmanyfields %>
            <td><b><%=rstemp(i).name%></B></TD>
    <% next %>
    </tr>
    <% ' Now lets grab all the records
    do while not rstemp.eof %>
        <tr>
        <% for i = 0 to howmanyfields
            thisvalue=rstemp(i)
            If isnull(thisvalue) then
                thisvalue="&nbsp;"
            end if%>
            <td valign=top><%=thisvalue%></td>
        <% next %>
        </tr>
        <%rstemp.movenext
    loop%>
    </table>
    <%
    rstemp.close
    set rstemp=nothing
    conntemp.close
    set conntemp=nothing
end sub

sub query2count(inputquery,inputDSN)
Dim oConn
Set oConn = Server.CreateObject("ADODB.Connection") 
Dim oRS
Set oRS = Server.CreateObject("ADODB.Recordset") 

oConn.Open inputDSN

oRS.Open inputquery, oConn, 1, 3 

Response.Write(oRs.RecordCount)
oRs.Close
oConn.Close 
Set oConn = Nothing
end sub

sub createtable(tablename,inputDSN)

Dim tblNew

Set tblNew = Server.CreateObject("ADOX.Table") 

		tblNew.Name = tablename 

Dim col ' As ADOX.Column
		Set col = Server.CreateObject("ADOX.Column")
		With col
			.ParentCatalog = inputDSN
			.Type = 3 ' adInteger
			.Name = "ID"
			.Properties("Autoincrement") = True
		End With 
		tblNew.Columns.Append col
	catDB.Tables.Append tblNew
	  Set col = Nothing
	  Set tblNew = Nothing
	  Set catDB = Nothing	

end sub
%>
