<%
'<------- /같같같같같같같같같같같같같\ ------->
'<------ / Gerando com PageMaster 2.5 \ ------>
'<----- /   http://www.gnove.com.br    \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_CONFIGURACOES_opt.asp
'CRIADO EM.........: 17/6/2007 18:52:44
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "default.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - CONFIGURACOES"
ElseIf acao = "2" Then
   sNome = "SELECINAR MODELO DO PORTAL"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>CONFIGURACOES - PageMaster 2.0 :::.</title>
<meta name="generator" content="PageMaster v2.0"/>
<script language="JavaScript" src="validator.js"></script>
<script language=JavaScript src="popcalendar.js"></script>
<script language="Javascript1.2" src="editor.js"></script>
<script>_editor_url = "";</script>
<SCRIPT language=Javascript>

function go2url() {
history.go(-1);
}
function Validator(theForm)
{
if (theForm.cp_a24Template && !EW_hasValue(theForm.cp_a24Template, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_a24Template, "SELECT", "Selecione um valor para o campo LAYOUT"))
            return false;
    }
if (theForm.cp_a01corsite && !EW_hasValue(theForm.cp_a01corsite, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_a01corsite, "SELECT", "Selecione um valor para o campo COR DO SITE"))
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
<meta name="robots" content="none">
<link rel="stylesheet" href="pagemaster.css">
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000" leftmargin=0 topmargin=0>
<table width='100%' cellpadding='0' cellspacing='0' height='100%'>
    <tr>
        <td align='center' valign='middle'>
<table cellpadding="0" cellspacing="0" width="70%" align="center">
    <tr>
        <td " class="tabelatitulo" align="center" height="50">
            <p>-:- <%= sNome %> -:-</p>
        </td>
    </tr>
    <tr>
        <td height="5"" class="tabelaitem">
        </td>
    </tr>
    <tr>
        <td" align="center" valign="top">
            <table align="center" border="1" cellspacing="0" width="70%" bordercolordark="white" bordercolorlight="black">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table align="center" border=1 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_CONFIGURACOES_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">LAYOUT<br></span>
           <select name="cp_a24Template" size="1" class="ud_caixa">
<%sLin = "Select * From CORES"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option selected>Selecione</option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("a00id") %>"><%= objRs("a08NomeCor") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
       <TD><P><span class="aFontePT">COR DO SITE<br></span>
           <select name="cp_a01corsite" size="1" class="ud_caixa">
<%sLin = "Select * From CORES"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option selected>Selecione</option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("a00id") %>"><%= objRs("a08NomeCor") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
sNot = 1
sSel = "Select * From CONFIGURACOES Where a00idconfig =1"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a24Template             = Trim(objRS("a24Template"))
va_a01corsite              = Trim(objRS("a01corsite"))
'-----------------------------------------------
%>
                            <table align="center" border=1 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_CONFIGURACOES_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">LAYOUT DO PORTAL<br></span>
           <select name="cp_a24Template" size="1" class="ud_caixa">
<%
sLin = "Select * From CORES"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option value="<%=va_a24Template%>" selected><%=va_a24Template%></option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("a00id") %>"><%= objRs("a08NomeCor") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
       <TD><P><span class="aFontePT">COR DO PORTAL<br></span>
           <select name="cp_a01corsite" size="1" class="ud_caixa">
<%
sLin = "Select * From CORES"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option value="<%=va_a01corsite%>" selected><%=va_a01corsite%></option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("a00id") %>"><%= objRs("a08NomeCor") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = 1
sSel = "Delete From CONFIGURACOES Where a00idconfig =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_CONFIGURACOES.asp"
ElseIf acao = "4" Then
sNot = request.querystring("sF")
Response.Expires = 0
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest RequestBin
Dim sTipo, sCmd
sTipo  = request.querystring("tp")
'-----------------------------------------------
va_a24Template             = Trim(UploadRequest.Item("cp_a24Template").Item("Value"))
va_a01corsite              = Trim(UploadRequest.Item("cp_a01corsite").Item("Value"))
'-----------------------------------------------

If sTipo = "1" Then
sNot = 1
   sSel = sSel & "INSERT INTO CONFIGURACOES(a24Template, a01corsite)"
   sSel = sSel & "VALUES ('" & va_a24Template & "', '" & va_a01corsite & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "entrada.asp"

ElseIf sTipo = "2" Then

sNot = 1
   sSel = sSel & "UPDATE CONFIGURACOES SET "
   sSel = sSel & "a24Template = '" & va_a24Template & "', a01corsite = '" & va_a01corsite & "' "
   sSel = sSel & "WHERE a00idconfig = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "entrada.asp"
Else
   sSel = sSel & "UPDATE CONFIGURACOES SET "
   sSel = sSel & "a24Template = '" & va_a24Template & "', a01corsite = '" & va_a01corsite & "' "
   sSel = sSel & "WHERE a00idconfig = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "entrada.asp"
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
<table cellpadding="0" cellspacing="0" width="70%" align="center">
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

</td>
    </tr>
</table></body>
</html>
<%
Sub Errou
If err <> 0 Then
%>
<br>
<table cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
        <td align="center" valign="middle">
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" align="center" cellpadding="0" cellspacing="0" width="571" bordercolordark="black" bordercolorlight="black">
                <tr>
                    <td width="571" height="63" align="center">
                        <p style="margin-left:5pt;"><span class="titulo">Informa豫o do sistema...</span></p>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="128" align="center">
           <span class="titulo">Erro n? <%=err.Number%> <br><%=err.Source%><br><%= err.Description%></span>
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

