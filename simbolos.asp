<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">SÍMBOLOS DO MUNÍCIPIO DE <%= iCidade %></font></td>

<%
	Dim sFoto, iIdPas
	iIdPas = Request("id")
	sFoto = 0
%>
<table width="127" cellspacing="2" cellpadding="2" align="center">
    <tr>
<%
	SQL = "Select * From simbolos Order By a00id desc"
	Set rsGal = Server.CreateObject("ADODB.Recordset")
	rsGal.Open SQL, objConect, 3, 3
	Do While not rsGal.EOF
%>
        <td width="117" height="110" style="border-width:1px; border-style:outset;" align="center" valign="middle">
		<a href="#" onClick="ajax('simbolos_view.asp?id=<%= rsGal(0) %>','conteudo')"><img src="config/simbolos/<%= rsGal(1) %>" alt="<%= rsGal(2) %>" width="97" height="97" border="0"></a>
		</td>
<%
	sFoto = sFoto + 1
If sFoto = 4 then
	Response.Write "            </tr>" & vbCrLf
	Response.Write "            <tr>"
	sFoto = 0
End If

	rsGal.MoveNext
	Loop
	rsGal.close
%>
    </tr>
</table>
    </tr>
</table>