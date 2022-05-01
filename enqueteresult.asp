<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
	<link rel="StyleSheet" HREF="padroes/enquete.css" type="text/css">
	<title>Enquete - Resultado</title>
<style>
<!--
@import url(suport/suport.css);
-->
</style>
<SCRIPT language="JavaScript">
function voltar() {
	window.close();
}
</SCRIPT>
<meta name="robots" content="none">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#FF0000" leftmargin="0" topmargin="0" marginwidth="0">

<table align="center" cellpadding="0" cellspacing="0" width="314">
    <tr>
        <td width="314" height="50" align="center" bgcolor="<%= cCorTopoBase %>"><strong><font size="2" face="Verdana,'Trebuchet MS'" color="<%= ca05corLink %>"><%= iTitulo %></font></strong></td>
    </tr>
    <tr>
        <td width="315" align="center"><strong><font size="2" face="Verdana,'Trebuchet MS'">ENQUETE</font></strong></td>
    </tr>
    <tr>
        <td align="center" width="315" valign="middle">
            <p><%
Dim XEnquete
	lAnketID = request("id")
	lSecenekID = request("optSecenek")
	UserIP = Request.ServerVariables("REMOTE_ADDR")
	
	XEnquete = Request.Cookies("Enquete")("Cod")
	
	if lSecenekID = 0 then
		%>
</p>
            <table align="center" border="0" cellpadding="0" cellspacing="0" width="95%">
                <tr>
                    <td><br><br>
                        <p class="smalltext" align="center">
			<b><font face="Verdana" size="2" color="#FF0000">Selecione um item para votar.<br></font></b></p>
                    </td>
                </tr>
                <tr>
                    <td height="51">
                        <p class="smalltext" align="center">
	<a href="#" onClick="history.go(-1)"><img src="imagens/botao_anterior.gif" alt="" border="0"></a>
<%
		response.end
	end if

	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	set rsTemp.ActiveConnection = objConect
	rsTemp.CursorType = adOpenStatic

	sSQL = "SELECT IP FROM enquetevotos WHERE SurveyID = " & lAnketID
	rsTemp.Open sSQL, objConect, 3 ,3
	
	If XEnquete <> lAnketID  then
	
		rsTemp.close
		sSQL = "SELECT SurveyID, OptionID, IP FROM enquetevotos WHERE AnswerID = 0"
		rsTemp.Open sSQL, objConect, 3 ,3
	
		rsTemp.AddNew
		rsTemp("SurveyID") = lAnketID
		rsTemp("OptionID") = lSecenekID
		rsTemp("IP") = UserIP
		
		rsTemp.update
		rsTemp.close
		
		Response.Cookies("Enquete")("Cod") = lAnketID
		Response.Cookies("Enquete")("Data")= formatdatetime(now(),vbshortdate)
		Response.Cookies("Enquete").Expires = dateadd("d",40,now())
				
		%></p>
                    </td>
                </tr>
            </table>
            <p class="smalltext"></p>
            <table border="0" cellspacing="0" cellpadding="0" width="95%" align="center">
                <tr>
                    <td colspan=2 class="header2" height="16">

                        <p align="center"><b><font face="Verdana" size="2" color="#000080">Obrigado por participar!</font></b><hr></td>
                </tr>
<%
	sSQL = "SELECT COUNT(AnswerID) AS TotalAnswers From enquetevotos WHERE SurveyID = " & lAnketID
	rsTemp.Open sSQL, objConect, 3 ,3
	lToplamOy = rsTemp("TotalAnswers")
	rsTemp.close

	sSQL = "SELECT Count(enquetevotos.OptionID) AS TotalAnswers, enquetesperguntas.OptionText " 
	sSQL = sSQL & "FROM enquetevotos INNER JOIN enquetesperguntas ON enquetevotos.OptionID = enquetesperguntas.OptionID "
	sSQL = sSQL & "WHERE enquetevotos.SurveyID=" & lAnketID &  " "
	sSQL = sSQL & "GROUP BY enquetesperguntas.OptionText "
	sSQL = sSQL & "ORDER BY Count(enquetevotos.OptionID) DESC;"

	rsTemp.Open sSQL, objConect, 3 ,3
	
	response.write "<tr><td colspan=2><font size=1 face=Verdana><b>Total " & lToplamOy & " votos.</b><br></td></tr>"
	
	do while not rsTemp.eof
		if lToplamOy <> 0 then
			lYuzde = (rsTemp("TotalAnswers") / lToplamOy) * 100
		else
			lYuzde = 0
		end if
		
		%>
                <tr>
                    <td bgcolor= "#ffffff"><font size="1" face="Verdana" color="000000"><b><%=rsTemp("OptionText")%></b></font></td>
                    <td bgcolor= "#ffffff" align=right><font size="1" face="Verdana" color="000000"><b><%=formatnumber((rsTemp("TotalAnswers")/lToplamOy)*100,2) & "%"%></b></font></td>
                </tr>
                <tr>
                    <td bgcolor= "#ffffff" colspan=2><img src="imagens/pb.gif" height=8 width=<%=(lYuzde)*0.9%>%"><img src="imagens/pbw.gif" height=8 width=<%=(100-lYuzde)*0.9%>%"></td>
                </tr>
                <%
		rsTemp.movenext
	loop
%>
            </table>
            <p><%Else%>
</p>
            <table border="0" cellspacing="0" cellpadding="0" width="95%" bgcolor="#ffffff" align="center">
                <tr>
                    <td colspan=2 bgcolor="#ffffff" height="16"><br><br>
                        <p align="center"><b><font face="Verdana" size="2" color="#000080">Seu voto já foi computado!!</font></b>
						<hr>
						</td>
                </tr>
<%
	rsTemp.Close
	sSQL = "SELECT COUNT(AnswerID) AS TotalAnswers From enquetevotos WHERE SurveyID = " & lAnketID
	rsTemp.Open sSQL, objConect, 3 ,3
	lToplamOy = rsTemp("TotalAnswers")
	rsTemp.close

	sSQL = 		  "SELECT Count(enquetevotos.OptionID) AS TotalAnswers, enquetesperguntas.OptionText " 
	sSQL = sSQL & "FROM enquetevotos INNER JOIN enquetesperguntas ON enquetevotos.OptionID = enquetesperguntas.OptionID "
	sSQL = sSQL & "WHERE enquetevotos.SurveyID=" & lAnketID &  " "
	sSQL = sSQL & "GROUP BY enquetesperguntas.OptionText "
	sSQL = sSQL & "ORDER BY Count(enquetevotos.OptionID) DESC;"

	rsTemp.Open sSQL, objConect, 3 ,3

	response.write "<tr><td bgcolor=""#ffffff"" colspan=2><font size=1 face=Verdana color=""#000000""><b>Total " & lToplamOy & " votos.</b></td></tr>"

	do while not rsTemp.eof
		if lToplamOy <> 0 then
			lYuzde = (rsTemp("TotalAnswers") / lToplamOy) * 100
		else
			lYuzde = 0
		end if
		
		%>			
                <tr>
                    <td bgcolor= "#ffffff"><font size="1" face="Verdana" color="" #ffffff""><b><%=rsTemp("OptionText")%></b></font></td>
                    <td bgcolor= "#ffffff" align=right><font size="1" face="Verdana" color="#000000"><b><%=formatnumber((rsTemp("TotalAnswers")/lToplamOy)*100,2) & "%"%></b></font></td>
                </tr>
                <tr>
                    <td bgcolor= "#ffffff" class="smalltext" colspan=2><img src="imagens/pb.gif" height=8 width=<%=(lYuzde)*0.9%>%"><img src="imagens/pbw.gif" height=8 width=<%=(100-lYuzde)*0.9%>%"></td>
                </tr>
                <%
		rsTemp.movenext
	loop
%>
            </table>
            <p><%End If
Set objConect = Nothing
%>
</p>
            <p class="smalltext" align="center">
<a href="javascript:voltar()"><img src="imagens/botao_fechar.gif" alt="" border="0"></a></p>
        </td>
    </tr>
</table>
</body>
</html>