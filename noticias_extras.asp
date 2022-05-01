<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">NOTÍCIAS</font></td>
							</tr>
<%
Dim SQL1
SQL1 = "Select a00id, a01data, a02titulo From NOTICIAS Order By a01data Desc"
    Set rsTempG = Server.CreateObject("ADODB.Recordset")
    rsTempG.Open SQL1, objConect, 3, 3
	Do While Not rsTempG.EOF
%>
    <tr>
        <td height="20"><img src="imagens\marcador.gif" width="14" height="19" border="0"></td>
        <td height="25" valign="top" class="navtext"><a href="#" onClick="ajax('noticias_view.asp?id=<%= rsTempG(0) %>','conteudo')"><%= Left(UCase(rsTempG(2)),56) %>...</a></td>
    </tr>
<%
rsTempG.MoveNext
Loop
rsTempG.Close
Set rsTempG = Nothing
%>
    </table>
	<table>
    <tr>
        <td height="35" colspan="5" class="navtext"><strong><%= date %></strong></td>
	</tr>
    <tr>
        <td colspan="5">
<!--#include file="rss_reader.asp" -->
<%
'Le arquivo RSS
pagina  = "http://www.estadao.com.br/rss/nacional.xml"
	Set RSS = new kwRSS_reader
	rssURL = pagina
	RSS.ParseLocation(rssURL)

	while not RSS.EOF
%>
		<b><font color="#0000FF">»</font></b>&nbsp;<a href="javascript:open_window('win', 'externa.asp?xQual=<% =RSS.GetLink %>', 0, 0, 800, 600, 0, 0, 0, 1, 1);" title="<% =Server.HTMLEncode(RSS.GetDesc) %>"><span class="navtext"><% =RSS.GetTitle %></span></a><br />
<% 
		RSS.MoveNext
	wend

if RSS.GetTextInput <> "" then
	Response.Write "</td></tr><tr><td align=""center"">"
	Response.Write RSS.GetTextInput
end if

if RSS.GetImage <> "" then
	Response.Write "</td></tr><tr><td align=""right"">"
	Response.Write RSS.GetImage
end if

Set RSS = nothing
%>
</td>
    </tr>
</table>
<br>
<a href="http://www.brasil.gov.br/" target="_blank"><img src="imagens/brasil_logo.jpg" alt="" width="400" height="40" border="0"></a>

