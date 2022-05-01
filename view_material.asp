<%Response.ContentType = "text/html; charset=iso-8859-1"%>

<!--#include file="_cnx.asp"-->
<%
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.addHeader "pragma","no-cache"
	Response.addHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	Session.lCid = 1033

sIDs = Request("id")
sCat = Request("cat")
sSub = Request("sub")

SQL = "Select * From CATEGORIAS Where A00IDC1=" & sCat
    Set rsTempC = Server.CreateObject("ADODB.Recordset")
    rsTempC.Open SQL, objConect, 3, 3

SQL2 = "Select * From SUBCATEG Where A00IDSUB=" & sSub
    Set rsTempS= Server.CreateObject("ADODB.Recordset")
    rsTempS.Open SQL2, objConect, 3, 3

Response.Write "<br><p class=""navtext"">Você esta em:<br><font color='#000000'><a href='default.asp'>HOME</a> - <a href='#' onClick=" & chr(34) & "ajax('verifica.asp?id="& sCat &"','conteudo')" & chr(34) & ">" & rsTempC(1) & "</a> - <a href='#'  onClick=" & chr(34) & "ajax('view.asp?sub="& sSub &"&cat="& sCat &"','conteudo')" & chr(34) & ">" & UCase(rsTempS(2)) & "</a></font><hr size='1' noshade></p>"

If sIDs <> "" Then
SQL2 = "Select * From MATERIAL Where A00ID=" & sIDs
    Set rsTempH = Server.CreateObject("ADODB.Recordset")
    rsTempH.Open SQL2, objConect, 3, 3

If rsTempH(8) = "1" Then
	If Session("iUser") = True Then
%>
		<p align="left"><a href="#" onClick="ajax('view.asp?sub=<%= sSub %>&cat=<%= sCat %>','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a></p>
		<span class="navtext"><font color="#000000"><%=rsTempH(1)%></font></span><br><br>
		<span class="navtext"><p align="justify"><%= rsTempH(2)%></p></span>
<% If rsTempH(10)<>"" Then %>
		<span class="navtext"><p>Clique no ícone para baixar o arquivo &raquo; <a href="config/arquivos/<%= rsTempH(10) %>" target="_blank"><img src="imagens/download_icon.gif" alt="Baixar arquivo!!" border="0" align="absmiddle"></a></p></span>
<% End If %>
<%
	Else
		Response.Write "<p align='center'><span class='tittext'><strong><font size='3' color='#FF0000'>Para ter acesso a este material você precisa ser cadastrado e/ou estar logado no portal!</font></strong></span></p>"
		Response.Write "<p align='center'><a href='#' onClick="& chr(34) & "ajax('assine_ja.asp','conteudo')"& chr(34) &"><img src='imagens/assine_ja.gif' border='0'></a></p><br><br><br><br><br>"
	End If
Else
%>
		<p align="left"><a href="#" onClick="ajax('view.asp?sub=<%= sSub %>&cat=<%= sCat %>','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a></p>
		<span class="tittext"><font color="#000000"><%=rsTempH(1)%></font></span><br><br>
		<span class="navtext"><p align="justify"><%= rsTempH(2)%></p></span>
<% If rsTempH(3)<>"" Then %>
		<span class="navtext"><p align="justify"><%= rsTempH(3)%></p></span>
<% End If %>
<% If rsTempH(10)<>"" Then %>
		<span class="navtext"><p>Clique no ícone para baixar o arquivo &raquo; <a href="config/arquivos/<%= rsTempH(10) %>" target="_blank"><img src="imagens/download_icon.gif" alt="Baixar arquivo!!" border="0" align="absmiddle"></a></p></span>
<% End If %>
<%
End If
rsTempH.Close
Set rsTempH = Nothing
End If
%>
<p align="left"><a href="#" onClick="ajax('view.asp?sub=<%= sSub %>&cat=<%= sCat %>','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0" align="middle"></a></p>
<hr>