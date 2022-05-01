<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<%
sIDs = Request("id")
If request("i")="0" Then %>
<p align="right"><a href="#" onClick="window.print()"><img src="imagens/botao_imprimir.gif" alt="" border="0"></a></p>
<link type="text/css" rel="stylesheet" href="suport/suport_def.css">
<% End If%>
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">NOTÍCIAS</font></td>
							</tr>
						</table>
<% If request("i")="0" Then %>
<span class="bodytext"><%= iTitulo & " " & iLegis &"</span><hr>"%>
<% End If%>
<%
If sIDs <> "" Then
SQL2 = "Select * From NOTICIAS Where a00id=" & sIDs
    Set rsTempH = Server.CreateObject("ADODB.Recordset")
    rsTempH.Open SQL2, objConect, 3, 3
%>
<br>
<span class="tittext"><strong><%=rsTempH(2)%></strong></span><br><br>
<span class="bodytext"><p align="justify"><% If rsTempH(5)<> "" Then %><img src="config/noticias/<%= rsTempH(5) %>" border="0" align="left"><% End If %><%= rsTempH(3)%></p><br>
Fonte: <%=rsTempH(4)%></span>
<%
rsTempH.Close
Set rsTempH = Nothing
End If
%>
<p align="center"><% If request("i")<>"0" Then %><a href="#" onClick="ajax('noticias_extras.asp','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a>&nbsp;<a href="javascript:open_window('imprimir', 'noticias_view.asp?id=<%= sIDs %>&i=0', 5, 5, 640, 480, 0, 0, 1, 1, 0)"><img src="imagens/botao_imprimir.gif" alt="" border="0"></a><% End If %></p>
<hr>