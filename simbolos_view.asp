<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">SÍMBOLOS DO MUNÍCIPIO DE <%= iCidade %></font></td>
<%
	Dim iIdPas
	iIdPas = Request("id")
	SQL = "Select * From simbolos Where a00id="& iIdPas
	Set rsGal = Server.CreateObject("ADODB.Recordset")
	rsGal.Open SQL, objConect, 3, 3
%>
<table width="99%" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td align="center"><br><img src="config/simbolos/<%= rsGal(1) %>" alt="<%= rsGal(2) %>" border="0" /></td>
    </tr>
    <tr>
        <td align="center" class="navtext"><%= rsGal(2) %></td>
    </tr>
    <tr>
        <td align="center"><br><a href="#" onClick="ajax('simbolos.asp','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a></td>
    </tr>
</table>
    </tr>
</table>
<%
	rsGal.close
%>