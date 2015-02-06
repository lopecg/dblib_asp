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
end sub%>
