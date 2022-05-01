<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_enquetes_opt.asp
'CRIADO EM.........: 16/2/2005 17:43:49
'-------------------------------------------------------------------------------
'On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "Novo Registro"
ElseIf acao = "2" Then
   sNome = "Editar Registro"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>.::: PageMaster 2.0 :::.</title>
<meta name="generator" content="PageMaster v2.0"/>
<script language="JavaScript" src="lu.js"></script>
<script language="JavaScript" src="validator.js"></script>
<SCRIPT language=Javascript>

function go2url() {
window.location = "pm_enquetes.asp";
}
function Validator(theForm)
{
if (theForm.cp_SurveyName && !EW_hasValue(theForm.cp_SurveyName, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_SurveyName, "TEXT", "Informe um valor para o campo NAME"))
            return false;
    }
if (theForm.cp_SurveyQuestion && !EW_hasValue(theForm.cp_SurveyQuestion, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_SurveyQuestion, "TEXT", "Informe um valor para o campo QUESTAO"))
            return false;
    }
  if (theForm.cp_StartDate && !EW_hasValue(theForm.cp_StartDate, "TEXT" )) {
          if (!EW_onError(theForm, theForm.cp_StartDate, "TEXT", "Informe a data correta para o campo INICIO"))
              return false;
      }
  if (theForm.cp_StartDate && !EW_checkeurodate(theForm.cp_StartDate.value)) {
      if (!EW_onError(theForm, theForm.cp_StartDate, "TEXT", "Data inválida para o campo INICIO"))
          return false;
      }
  if (theForm.cp_FinishDate && !EW_hasValue(theForm.cp_FinishDate, "TEXT" )) {
          if (!EW_onError(theForm, theForm.cp_FinishDate, "TEXT", "Informe a data correta para o campo FIM"))
              return false;
      }
  if (theForm.cp_FinishDate && !EW_checkeurodate(theForm.cp_FinishDate.value)) {
      if (!EW_onError(theForm, theForm.cp_FinishDate, "TEXT", "Data inválida para o campo FIM"))
          return false;
      }
if (theForm.cp_Active && !EW_hasValue(theForm.cp_Active, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_Active, "SELECT", "Selecione um valor para o campo ATIVO"))
            return false;
    }
if (theForm.cp_Active && !EW_hasValue(theForm.cp_Active, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_Active, "TEXT", "Informe um valor para o campo ATIVO"))
            return false;
    }
  return true;
}
function fotoZoom(img){
  foto1= new Image();
  foto1.src=(img);
  Controlla(img);
}
function Controlla(img){
  if((foto1.width!=0)&&(foto1.height!=0)){
    viewFoto(img);
  }
  else{
    funzione="Controlla('"+img+"')";
    intervallo=setTimeout(funzione,20);
  }
}
function viewFoto(img){
  largh=foto1.width+20;
  altez=foto1.height+20;
  stringa="width="+largh+",height="+altez;
  finestra=window.open(img,"",stringa);
}
</SCRIPT>
<script language=JavaScript src="popcalendar.js"></script>
<meta name="robots" content="none">
<link rel="stylesheet" href="pagemaster.css">
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000" leftmargin=0 topmargin=0>
<table cellpadding="0" cellspacing="0" width="610" align="center">
    <tr>
        <td " class="tabelatitulo" align="center" height="50">
            <p><span class="aFonte">-:- <%= sNome %> -:-</span></p>
        </td>
    </tr>
    <tr>
        <td height="5"" class="tabelaitem">
        </td>
    </tr>
    <tr>
        <td" align="center" valign="top">
            <table align="center" border="1" cellspacing="0" width="605" bordercolordark="white" bordercolorlight="#808080">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table align="center" border=1 cellspacing=0 width="605" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_enquetes_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">NAME<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_SurveyName></P></TD>
       <TD><P><span class="aFontePT">QUESTAO<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_SurveyQuestion></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">DATA DO INICIO<br></span><input type="text" name="cp_StartDate" value="<%=date%>" class="ud_caixa" style="WIDTH:150px">&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Selecione a data" onClick="popUpCalendar(this, this.form.cp_StartDate,'dd/mm/yyyy');return false;">&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=7', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a></P></TD>
       <TD><P><span class="aFontePT">DATA DO FIM<br></span><input type="text" name="cp_FinishDate" value="<%=dateadd("d",15,date)%>" class="ud_caixa" style="WIDTH:150px">&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Selecione a data" onClick="popUpCalendar(this, this.form.cp_FinishDate,'dd/mm/yyyy');return false;">&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=8', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a></P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">ENQUETE ATIVA<br></span>
                    <input type="radio" name="cp_Active" value="1"><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_Active" value="0" checked><span class="fontTD">NÃO</span>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=2', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
       </td>
</TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then %>
<%
On Error Resume Next
sNot = request.querystring("sF")
sSel = "Select * From enquetes Where SurveyID =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_SurveyName              = Trim(objRS("SurveyName"))
va_SurveyQuestion          = Trim(objRS("SurveyQuestion"))
va_StartDate               = Trim(objRS("StartDate"))
va_FinishDate              = Trim(objRS("FinishDate"))
va_Active                  = Trim(objRS("Active"))
%>
                            <table align="center" border=1 cellspacing=0 width="605" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_enquetes_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">NAME<br></span><INPUT class=ud_caixa style="WIDTH:295px" value='<% = va_SurveyName %>' name=cp_SurveyName></P></TD>
       <TD><P><span class="aFontePT">QUESTAO<br></span><INPUT class=ud_caixa style="WIDTH:295px" value='<% = va_SurveyQuestion %>' name=cp_SurveyQuestion></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">DATA DE INICIO<br></span><INPUT class=ud_caixa style="WIDTH:150px" value='<% = va_StartDate %>' name=cp_StartDate>&nbsp;<input type="image" src="images/pm_calendar.gif"  align="middle" alt="Selecione a data" onClick="popUpCalendar(this, this.form.cp_StartDate,'dd/mm/yyyy');return false;">&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=7', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a></P></TD>
       <TD><P><span class="aFontePT">DATA DO FIM<br></span><INPUT class=ud_caixa style="WIDTH:150px" value='<% = va_FinishDate %>' name=cp_FinishDate>&nbsp;<input type="image" src="images/pm_calendar.gif"  align="middle" alt="Selecione a data" onClick="popUpCalendar(this, this.form.cp_FinishDate,'dd/mm/yyyy');return false;">&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=8', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a></P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">ENQUETE ATIVA<br></span>
                    <input type="radio" name="cp_Active" value="1"<% If va_Active = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_Active" value="0"<% If va_Active = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=5', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
		</TD>
</TR>
                        </form>
<% '--------------------------------------------fim itens da enquete
if sNot <> 0 then
%>
                            </table>
	<table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="#DEDFEF">
		
		<tr>
			<td colspan=5 class="header1" align=center bgcolor="#DEDFEF" height="18">
			<font face="Verdana" size="2"><b>Opções da Enquete</b></font></td>
		</tr>
		
		<tr>
			<td width="25" bgcolor="white">&nbsp;</td>
			<td width="15" bgcolor="white">&nbsp;</td>
			<td class="smalltext" bgcolor="white">
			
			<a href="pm_enquetesperguntas_opt.asp?sF=<%=sNot%>&acao=1"><span class="aFontePT">(Adcionar Nova Opção)</span></a></td>
			<td bgcolor="white">&nbsp;</td>
			<td bgcolor="white">&nbsp;</td>
		</tr>
<%
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	Set rsTemp.ActiveConnection = objConect
	rsTemp.CursorType = adOpenStatic
	sSQL = "SELECT COUNT(AnswerID) AS TotalAnswers From enquetevotos WHERE SurveyID = " & sNot
	rsTemp.Open sSQL, ,3,3
	lToplamOy = rsTemp("TotalAnswers")
	rsTemp.close

	sSQL = "SELECT Count(enquetevotos.OptionID) AS TotalAnswers, enquetesperguntas.OptionID, enquetesperguntas.OptionText " 
	sSQL = sSQL & "FROM enquetesperguntas LEFT JOIN enquetevotos ON enquetesperguntas.OptionID = enquetevotos.OptionID "
	sSQL = sSQL & "WHERE enquetesperguntas.SurveyID=" & sNot &  " " 
	sSQL = sSQL & "GROUP BY enquetesperguntas.OptionID, enquetesperguntas.OptionText "
	sSQL = sSQL & "ORDER BY Count(enquetevotos.OptionID) DESC, enquetesperguntas.OptionID;"

	rsTemp.Open sSQL, ,3,3
	
	response.write "<tr><td></td><td></td><td class=""aFontePT""><b>Total de " & lToplamOy & " votos.</b></td><td></td><td></td></tr>"
	
	do while not rsTemp.eof
		if lToplamOy <> 0 then
			lYuzde = (rsTemp("TotalAnswers") / lToplamOy) * 100
		else
			lYuzde = 0
		end if
		lSecenek = lSecenek + 1
		%>
			<tr>
				<td class="smalltext" bgcolor="white"><a href='pm_enquetesperguntas_opt.asp?acao=3&sF=<%= sNot %>&sQ=<%=rsTemp("OptionID")%>'><span class="aFontePT">(Apagar)</span></a></td>
				<td class="header2" align=center bgcolor="white"><span class="aFontePT"><%=lSecenek%>.</span></td>
				<td class="smalltext" bgcolor="white"><a href='pm_enquetesperguntas_opt.asp?acao=2&sF=<%= sNot %>&sQ=<%=rsTemp("OptionID")%>'><span class="aFontePT"><%=rsTemp("OptionText")%></span></a></td>
				<td class="smalltext" align=right bgcolor="white"><span class="aFontePT"><%=formatnumber(lYuzde,2) & "%"%></span></td>
				<td width=100 bgcolor="white"><img src="images/pb.gif" height=8 width=<%=(lYuzde)*0.9%>%"><img src="images/pbw.gif" height=8 width=<%=(100-lYuzde)*0.9%>%"></td>
			</tr>
		<%
		rsTemp.movenext
	loop
end if
'--------------------------------------------fim itens da enquete
%>

<% ElseIf acao = "3" Then %>
<%
'sNot = request.querystring("sF")
'sSel = "Delete From enquetes Where SurveyID =" & sNot
'Set objRS = objConect.Execute(sSel)
'response.redirect "pm_enquetes.asp"

	lSurveyID = CLng(request("sF"))
	
	objConect.execute "DELETE FROM enquetes WHERE SurveyID = " & lSurveyID
	objConect.execute "DELETE FROM enquetesperguntas WHERE SurveyID = " & lSurveyID
	objConect.execute "DELETE FROM enquetevotos WHERE SurveyID = " & lSurveyID
	response.redirect request("http_referer")

ElseIf acao = "4" Then
sNot = request.querystring("sF")
Response.Expires = 0
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest RequestBin
Dim sTipo, sCmd
sTipo  = request.querystring("tp")
%>
<%
'-----------------------------------------------
va_SurveyName              = Trim(UploadRequest.Item("cp_SurveyName").Item("Value"))
va_SurveyQuestion          = Trim(UploadRequest.Item("cp_SurveyQuestion").Item("Value"))
va_StartDate               = Trim(UploadRequest.Item("cp_StartDate").Item("Value"))
va_FinishDate              = Trim(UploadRequest.Item("cp_FinishDate").Item("Value"))
va_Active                  = Trim(UploadRequest.Item("cp_Active").Item("Value"))

If sTipo = "1" Then
sNot = request.querystring("sF")
   sSel = sSel & "INSERT INTO enquetes(SurveyName, SurveyQuestion, StartDate, FinishDate, Active)"
   sSel = sSel & "VALUES ('" & va_SurveyName & "', '" & va_SurveyQuestion & "', '" & va_StartDate & "', '" & va_FinishDate & "', " & va_Active & ")"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_enquetes.asp"

ElseIf sTipo = "2" Then

'***********************************************
sNot = request.querystring("sF")
   sSel = sSel & "UPDATE enquetes SET "
   sSel = sSel & "SurveyName = '" & va_SurveyName & "', SurveyQuestion = '" & va_SurveyQuestion & "', StartDate = '" & va_StartDate & "', FinishDate = '" & va_FinishDate & "', Active = " & va_Active & " "
   sSel = sSel & "WHERE SurveyID = " & sNot
   response.write sSel
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_enquetes.asp"
Else
   sSel = sSel & "UPDATE enquetes SET "
   sSel = sSel & "SurveyName = '" & va_SurveyName & "', SurveyQuestion = '" & va_SurveyQuestion & "', StartDate = '" & va_StartDate & "', FinishDate = '" & va_FinishDate & "', Active = " & va_Active & " "
   sSel = sSel & "WHERE SurveyID = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_enquetes.asp"
End If
End If

Sub BuildUploadRequest(RequestBin)
PosBeg = 1
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(13)))
boundary = MidB(RequestBin, PosBeg, PosEnd - PosBeg)
BoundaryPos = InStrB(1, RequestBin, boundary)
Do Until (BoundaryPos = InStrB(RequestBin, boundary & getByteString("--")))
Dim UploadControl
Set UploadControl = CreateObject("Scripting.Dictionary")
Pos = InStrB(BoundaryPos, RequestBin, getByteString("Content-Disposition"))
Pos = InStrB(Pos, RequestBin, getByteString("name="))
PosBeg = Pos + 6
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(34)))
Name = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
PosFile = InStrB(BoundaryPos, RequestBin, getByteString("filename="))
PosBound = InStrB(PosEnd, RequestBin, boundary)
If PosFile <> 0 And (PosFile < PosBound) Then
PosBeg = PosFile + 10
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(34)))
FileName = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
UploadControl.Add "FileName", FileName
Pos = InStrB(PosEnd, RequestBin, getByteString("Content-Type:"))
PosBeg = Pos + 14
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(13)))
ContentType = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
UploadControl.Add "ContentType", ContentType
PosBeg = PosEnd + 4
PosEnd = InStrB(PosBeg, RequestBin, boundary) - 2
Value = MidB(RequestBin, PosBeg, PosEnd - PosBeg)
Else
Pos = InStrB(Pos, RequestBin, getByteString(Chr(13)))
PosBeg = Pos + 4
PosEnd = InStrB(PosBeg, RequestBin, boundary) - 2
Value = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
End If
UploadControl.Add "Value", Value
UploadRequest.Add Name, UploadControl
BoundaryPos = InStrB(BoundaryPos + LenB(boundary), RequestBin, boundary)
Loop
End Sub
'===============================================
Function getByteString(StringStr)
For i = 1 To Len(StringStr)
Char = Mid(StringStr, i, 1)
getByteString = getByteString & ChrB(AscB(Char))
Next
End Function

Function getString(StringBin)
getString = ""
For intCount = 1 To LenB(StringBin)
getString = getString & Chr(AscB(MidB(StringBin, intCount, 1)))
Next
End Function
'===============================================
%>
            </td>
    </tr>
</table>
<table cellpadding="0" cellspacing="0" width="610" align="center">
    <tr>
        <td class="tabelaitem" height="5" colspan="2">
        </td>
    </tr>
    <tr>
        <td height="21" class="tabelatitulo" colspan="2">
        </td>
    </tr>
    <tr>
        <td></td>
        <td align="center" valign="top">
                    </td>
    </tr>
</table>
</body>
</html>
<%
Sub Errou
If err <> 0 Then
%>
<br>
<table cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
        <td align="center" valign="middle">
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" align="center" cellpadding="0" cellspacing="0" width="571" bordercolordark="black" bordercolorlight="#808080">
                <tr>
                    <td width="571" height="63" align="center">
                        <p style="margin-left:5pt;"><span class="titulo">Informação do sistema...</span></p>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="128" align="center">
           <span class="titulo">Erro nº <%=err.Number%> <br><%=err.Source%><br><%= err.Description%></span>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="38" align="center">
           <span class="titulo"><a href="javascript:self.history.go(-1)">&lt;&lt; Voltar</a></span>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<br>
<%
Response.End
End If
End Sub
%>

