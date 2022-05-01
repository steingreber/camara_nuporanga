<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_enquetesperguntas_opt.asp
'CRIADO EM.........: 17/2/2005 11:12:31
'-------------------------------------------------------------------------------

If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
On Error Resume Next
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "Nova opção da enquete " & request("sF")
ElseIf acao = "2" Then
   sNome = "Editar opção da enquete " & request("sF")
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>.::: PageMaster 2.0 :::.</title>
<meta name="generator" content="PageMaster v2.0"/>
<script language="JavaScript" src="lu.js"></script>
<SCRIPT language=Javascript>

function go2url() {
window.location = "pm_enquetes_opt.asp?acao=2&sF=<% =request("sF") %>";
}
function Validator(theForm)
{
if (theForm.cp_SurveyID && !EW_hasValue(theForm.cp_SurveyID, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_SurveyID, "TEXT", "Informe um valor para o campo ID ENQUETE"))
            return false;
    }
if (theForm.cp_OptionText && !EW_hasValue(theForm.cp_OptionText, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_OptionText, "TEXT", "Informe um valor para o campo OPÇÕES"))
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
                        <form action="pm_enquetesperguntas_opt.asp?acao=4&tp=1&sF=<%=Request("sF")%>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">NUMERO DA ENQUETE<br></span><input type="text" name="cp_SurveyID" value="<%=Request("sF")%>" readonly class="ud_caixa" style="WIDTH:100px"></P></TD>
       <TD><P><span class="aFontePT">OPÇÕES<br></span><INPUT class=ud_caixa style="WIDTH:395px" name=cp_OptionText></P></TD>
</TR>
<TR>
                        </form>    
                            </table>
<%
ElseIf acao = "2" Then

On Error Resume Next
sEnq = request("sQ")
sNot = request("sF")
sSel = "Select * From enquetesperguntas Where OptionID =" & sEnq
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_SurveyID                = Trim(objRS("SurveyID"))
va_SurveyID                = Replace(va_SurveyID,"<BR>", chr(13),1)
'-----------------------------------------------
va_OptionText              = Trim(objRS("OptionText"))
va_OptionText              = Replace(va_OptionText,"<BR>", chr(13),1)
%>
                            <table align="center" border=1 cellspacing=0 width="605" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_enquetesperguntas_opt.asp?acao=4&tp=2&sF=<% = sNot %>&sQ=<%= sEnq %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">NUMERO DA ENQUETE<br></span><input type="text" name="cp_SurveyID" value="<% = va_SurveyID %>" readonly class="ud_caixa" style="WIDTH:100px"></P></TD>
       <TD><P><span class="aFontePT">OPÇÕES<br></span><INPUT class=ud_caixa style="WIDTH:395px" value='<% = va_OptionText %>' name=cp_OptionText></P></TD>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sEnq = request("sQ")
sNot = request("sF")
sSel = "Delete From enquetesperguntas Where OptionID =" & sEnq
Set objRS = objConect.Execute(sSel)
response.redirect "pm_enquetes_opt.asp?acao=2&sF=" & sNot

ElseIf acao = "4" Then
sEnq = Request("sQ")
sNot = Request("sF")
sTipo  = Request("tp")
Response.Expires = 0
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest RequestBin
Dim sTipo, sCmd
'-----------------------------------------------
va_SurveyID                = Trim(UploadRequest.Item("cp_SurveyID").Item("Value"))
va_SurveyID                = Replace(va_SurveyID,chr(13),"<BR>" ,1)
'-----------------------------------------------
va_OptionText              = Trim(UploadRequest.Item("cp_OptionText").Item("Value"))
va_OptionText              = Replace(va_OptionText,chr(13),"<BR>" ,1)

If sTipo = "1" Then
   sSel = sSel & "INSERT INTO enquetesperguntas (SurveyID, OptionText)"
   sSel = sSel & "VALUES ('" & va_SurveyID & "', '" & va_OptionText & "')"

   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_enquetes_opt.asp?acao=2&sF=" & sNot

ElseIf sTipo = "2" Then
   sSel = sSel & "UPDATE enquetesperguntas SET "
   sSel = sSel & "SurveyID = '" & va_SurveyID & "', OptionText = '" & va_OptionText & "' "
   sSel = sSel & "WHERE OptionID = " & sEnq
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_enquetes_opt.asp?acao=2&sF=" & sNot
Else
   sSel = sSel & "UPDATE enquetesperguntas SET "
   sSel = sSel & "SurveyID = '" & va_SurveyID & "', OptionText = '" & va_OptionText & "' "
   sSel = sSel & "WHERE OptionID = " & sEnq
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_enquetes_opt.asp?acao=2&sF=" & sNot
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

