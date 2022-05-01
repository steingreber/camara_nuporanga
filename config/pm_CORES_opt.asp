<%
'<------- /같같같같같같같같같같같같같\ ------->
'<------ / Gerando com PageMaster 2.5 \ ------>
'<----- /   http://www.gnove.com.br    \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_CORES_opt.asp
'CRIADO EM.........: 16/6/2007 21:41:53
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "default.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - CORES"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - CORES"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>CORES - PageMaster 2.0 :::.</title>
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
if (theForm.cp_a08NomeCor && !EW_hasValue(theForm.cp_a08NomeCor, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a08NomeCor, "TEXT", "Informe um valor para o campo NOME DA COR"))
            return false;
    }
if (theForm.cp_a06Tamplate && !EW_hasValue(theForm.cp_a06Tamplate, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_a06Tamplate, "SELECT", "Selecione um valor para o campo LAYOUT"))
            return false;
    }
if (theForm.cp_a01corTopoBase && !EW_hasValue(theForm.cp_a01corTopoBase, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a01corTopoBase, "TEXT", "Informe um valor para o campo COR TOPO BASE"))
            return false;
    }
if (theForm.cp_a02corTopoTabela && !EW_hasValue(theForm.cp_a02corTopoTabela, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a02corTopoTabela, "TEXT", "Informe um valor para o campo COR TOPO TABELA"))
            return false;
    }
if (theForm.cp_a03corCorpoTabela && !EW_hasValue(theForm.cp_a03corCorpoTabela, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a03corCorpoTabela, "TEXT", "Informe um valor para o campo COR CORPO TABELA"))
            return false;
    }
if (theForm.cp_a04corFonte && !EW_hasValue(theForm.cp_a04corFonte, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a04corFonte, "TEXT", "Informe um valor para o campo COR FONTE"))
            return false;
    }
if (theForm.cp_a05corLink && !EW_hasValue(theForm.cp_a05corLink, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a05corLink, "TEXT", "Informe um valor para o campo COR LINK"))
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
            <p><span class="aFonte">-:- <%= sNome %> -:-</span></p>
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
                        <form action="pm_CORES_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">NOME DA COR<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_a08NomeCor></P></TD>
       <TD><P><span class="aFontePT">LAYOUT<br></span>
           <select name="cp_a06Tamplate" size="1" class="ud_caixa">
<%sLin = "Select * From TEMPLATES"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option selected>Selecione</option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("a00id") %>"><%= objRs("a01Template") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
</TR>
<TR>
       <TD><P><span class="aFontePT">COR TOPO BASE<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_a01corTopoBase></P></TD>
       <TD><P><span class="aFontePT">COR TOPO TABELA<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_a02corTopoTabela></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">COR CORPO TABELA<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_a03corCorpoTabela></P></TD>
       <TD><P><span class="aFontePT">COR FONTE<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_a04corFonte></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">COR LINK<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_a05corLink></P></TD>
</TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
sNot = request.querystring("sF")
sSel = "Select * From CORES Where a00id =" & sNot
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a08NomeCor              = Trim(objRS("a08NomeCor"))
va_a06Tamplate             = Trim(objRS("a06Tamplate"))
va_a01corTopoBase          = Trim(objRS("a01corTopoBase"))
va_a02corTopoTabela        = Trim(objRS("a02corTopoTabela"))
va_a03corCorpoTabela       = Trim(objRS("a03corCorpoTabela"))
va_a04corFonte             = Trim(objRS("a04corFonte"))
va_a05corLink              = Trim(objRS("a05corLink"))
'-----------------------------------------------
%>
                            <table align="center" border=1 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_CORES_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">NOME DA COR<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_a08NomeCor %>' name=cp_a08NomeCor></P></TD>
       <TD><P><span class="aFontePT">LAYOUT<br></span>
           <select name="cp_a06Tamplate" size="1" class="ud_caixa">
<%
sLin = "Select * From TEMPLATES"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option value="<%=va_a06Tamplate%>" selected><%=va_a06Tamplate%></option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("a00id") %>"><%= objRs("a01Template") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
</TR>
<TR>
       <TD><P><span class="aFontePT">COR TOPO BASE<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_a01corTopoBase %>' name=cp_a01corTopoBase></P></TD>
       <TD><P><span class="aFontePT">COR TOPO TABELA<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_a02corTopoTabela %>' name=cp_a02corTopoTabela></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">COR CORPO TABELA<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_a03corCorpoTabela %>' name=cp_a03corCorpoTabela></P></TD>
       <TD><P><span class="aFontePT">COR FONTE<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_a04corFonte %>' name=cp_a04corFonte></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">COR LINK<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_a05corLink %>' name=cp_a05corLink></P></TD>
</TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From CORES Where a00id =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_CORES.asp"
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
va_a08NomeCor              = Trim(UploadRequest.Item("cp_a08NomeCor").Item("Value"))
va_a06Tamplate             = Trim(UploadRequest.Item("cp_a06Tamplate").Item("Value"))
va_a01corTopoBase          = Trim(UploadRequest.Item("cp_a01corTopoBase").Item("Value"))
va_a02corTopoTabela        = Trim(UploadRequest.Item("cp_a02corTopoTabela").Item("Value"))
va_a03corCorpoTabela       = Trim(UploadRequest.Item("cp_a03corCorpoTabela").Item("Value"))
va_a04corFonte             = Trim(UploadRequest.Item("cp_a04corFonte").Item("Value"))
va_a05corLink              = Trim(UploadRequest.Item("cp_a05corLink").Item("Value"))
'-----------------------------------------------

If sTipo = "1" Then
sNot = request.querystring("sF")
   sSel = sSel & "INSERT INTO CORES(a08NomeCor, a06Tamplate, a01corTopoBase, a02corTopoTabela, a03corCorpoTabela, a04corFonte, a05corLink)"
   sSel = sSel & "VALUES ('" & va_a08NomeCor & "', '" & va_a06Tamplate & "', '" & va_a01corTopoBase & "', '" & va_a02corTopoTabela & "', '" & va_a03corCorpoTabela & "', '" & va_a04corFonte & "', '" & va_a05corLink & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CORES.asp"

ElseIf sTipo = "2" Then

sNot = request.querystring("sF")
   sSel = sSel & "UPDATE CORES SET "
   sSel = sSel & "a08NomeCor = '" & va_a08NomeCor & "', a06Tamplate = '" & va_a06Tamplate & "', a01corTopoBase = '" & va_a01corTopoBase & "', a02corTopoTabela = '" & va_a02corTopoTabela & "', a03corCorpoTabela = '" & va_a03corCorpoTabela & "', a04corFonte = '" & va_a04corFonte & "', a05corLink = '" & va_a05corLink & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CORES.asp"
Else
   sSel = sSel & "UPDATE CORES SET "
   sSel = sSel & "a08NomeCor = '" & va_a08NomeCor & "', a06Tamplate = '" & va_a06Tamplate & "', a01corTopoBase = '" & va_a01corTopoBase & "', a02corTopoTabela = '" & va_a02corTopoTabela & "', a03corCorpoTabela = '" & va_a03corCorpoTabela & "', a04corFonte = '" & va_a04corFonte & "', a05corLink = '" & va_a05corLink & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CORES.asp"
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
           <span class="titulo">Erro n� <%=err.Number%> <br><%=err.Source%><br><%= err.Description%></span>
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

