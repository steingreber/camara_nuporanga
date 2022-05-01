<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">FOTOS DO MUNÍCIPIO DE <%= iCidade %></font></td>
							</tr>
<%
	Dim iIdPas
	iIdPas = Request("id")
	SQL = "Select * From galeria Where a00id="& iIdPas
	Set rsGal = Server.CreateObject("ADODB.Recordset")
	rsGal.Open SQL, objConect, 3, 3
%>
<table width="99%" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td align="center" class="navtext"><%= rsGal(2) %></td>
    </tr>
    <tr>
        <td align="center"><img src="config/galeria/<%= rsGal(1) %>" alt="<%= rsGal(2) %>" border="0" /></td>
    </tr>
    <tr>
        <td align="center" class="navtext"><%= rsGal(2) %></td>
    </tr>
    <tr>
        <td align="center"><br><a href="#" onClick="ajax('galeria.asp','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a></td>
    </tr>
</table>
<%
	rsGal.close
%>