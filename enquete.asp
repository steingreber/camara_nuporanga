<html>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<head>
<title>Enquete</title>
<meta name="author" content="Alex L. Steingreber">
<meta name="robots" content="none">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<table width="314" cellpadding="0" cellspacing="0" height="100%">
    <tr>
        <td width="314" height="50" align="center" bgcolor="<%= cCorTopoBase %>"><strong><font face="Verdana,'Trebuchet MS'" size="2" color="<%= ca05corLink %>"><%= iTitulo %></font></strong></td>
    </tr>
    <tr>
        <td align="center" valign="top" height="20"><strong><font size="2" face="Verdana,'Trebuchet MS'">ENQUETE</font></strong></td>
    </tr>
    <tr>
        <td align="center" valign="top">
<%
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	set rsTemp.ActiveConnection = objConect
	'rsTemp.CursTemporType = adOpenStatic

	sSQL = "SELECT SurveyID, SurveyName, SurveyQuestion FROM enquetes WHERE SurveyID=" & Request.QueryString("eId")'& lAnketID
	rsTemp.Open sSQL, objConect, 3 ,3
	rsTemp.MoveLast
	lAnketID = rsTemp("SurveyID")

	if rsTemp.eof or rsTemp.bof then
		response.write "<font face=""Verdana"" color=""#FF0000"">Enquete não disponível!!</font>"
		response.end
	end if
	
	sSurveyName     = rsTemp("SurveyName")
	sSurveyQuestion = rsTemp("SurveyQuestion")
	
	rsTemp.close
%>
<form action="enqueteresult.asp?ID=<%=lAnketID%>" method="post">
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
	<tr>
		<td colspan="2" width="313"><br>
		<p align="center"><font size="2" face="Verdana"><b><%=sSurveyQuestion%></b></font></td>
	</tr>
	<%
	sSQL = "SELECT OptionID, OptionText, AdditionalInformation FROM enquetesperguntas WHERE SurveyID = " & lAnketID
	rsTemp.Open sSQL, objConect, 3 ,3
	do while not rsTemp.eof
	%>
		<tr>
			<td width="24" height="32">
				<input type="radio" name="optSecenek" value="<%=rsTemp("OptionID")%>">
			</td>
			<td width="289" height="32">
				<%If rsTemp("AdditionalInformation") <> "" then%>
					<a href="javascript:open_window('additionalinfo.asp?id=<%=rsTemp("OptionID")%>&name=<%=Server.URLEncode(rsTemp("OptionText"))%>')"><%=rsTemp("OptionText")%></a>
				<%Else%>
					<font size="1" face="Verdana"><%=rsTemp("OptionText")%></font>
				<%End If%>
			</td>
		</tr>
<%
	rsTemp.movenext
	loop
	rsTemp.close
%>
	<tr>
		<td align=center colspan="2" width="313" height="39">
				<input type="image" src="imagens/botao_votar.gif">
		</td>
	</tr>
</table>
			</form>
		
		</td>
    </tr>
    <tr>
        <td width="314" height="31" align="center" bgcolor="<%= cCorTopoBase %>"><a href="javascript:close()"><img src="imagens/botao_fechar.gif" alt="" border="0"></a></td>
    </tr>
</table>
</body>

</html>
