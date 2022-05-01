<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<table width="99%" cellpadding="1" cellspacing="1">
<%
iData = Request("dia") & "/" & Request("mes") & "/" & Request("ano")
SQL = "Select * From CALENDARIO Where o02data=#"& iData &"#"
    Set rsTempS= Server.CreateObject("ADODB.Recordset")
    rsTempS.Open SQL, objConect, 3, 3
	Do While not rsTempS.EOF
%>
    <tr>
        <td bgcolor="#ADC6E4">Data</td>
        <td bgcolor="#D5F0F9" class="data"><%= rsTempS(2) %></td>
    </tr>
    <tr>
        <td bgcolor="#ADC6E4">Titulo</td>
        <td bgcolor="#E7F0F9" class="data"><%= rsTempS(3) %></td>
    </tr>
    <tr>
        <td bgcolor="#ADC6E4">Descrição</td>
        <td bgcolor="#E7F0F9" class="data"><%= rsTempS(4) %></td>
    </tr>
    <tr>
        <td bgcolor="#ADC6E4">Documento</td>
        <td bgcolor="#E7F0F9" class="data"><%= rsTempS(5) %></td>
    </tr>
    <tr>
        <td bgcolor="#ADC6E4">Código</td>
        <td bgcolor="#E7F0F9" class="data"><%= rsTempS(6) %></td>
    </tr>
    <tr>
        <td height="5" colspan="2" bgcolor="#FFFFFF"></td>
    </tr>
<%
	rsTempS.MoveNext
	Loop
%>
</table>