<%
'<------- /같같같같같같같같같같같같같\ ------->
'<------ / Gerando com PageMaster 2.5 \ ------>
'<----- /   http://www.gnove.com.br    \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_CATEGORIAS_opt.asp
'CRIADO EM.........: 19/6/2007 23:01:50
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "default.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - CATEGORIAS"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - CATEGORIAS"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>CATEGORIAS - PageMaster 2.0 :::.</title>
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
if (theForm.cp_A01CATEG1 && !EW_hasValue(theForm.cp_A01CATEG1, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A01CATEG1, "TEXT", "Informe um valor para o campo CATEGORIA"))
            return false;
    }
  if (theForm.cp_A02EXIBIR && !EW_hasValue(theForm.cp_A02EXIBIR, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_A02EXIBIR, "RADIO", "Selecione um valor para EXIBIR!"))
            return false;
    }
  if (theForm.cp_A03HASUB && !EW_hasValue(theForm.cp_A03HASUB, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_A03HASUB, "RADIO", "Selecione um valor para HA SUB-CATEG!"))
            return false;
    }
if (theForm.cp_A03HASUB && !EW_hasValue(theForm.cp_A03HASUB, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A03HASUB, "TEXT", "Informe um valor para o campo A03HASUB"))
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
                        <form action="pm_CATEGORIAS_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">CATEGORIA<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A01CATEG1></P></TD>
       <TD><P><span class="aFontePT">EXIBIR&nbsp;</span>
                    <input type="radio" name="cp_A02EXIBIR" value="1"><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_A02EXIBIR" value="0"><span class="fontTD">N홒</span>&nbsp;
       </P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">HA SUB-CATEG&nbsp;</span>
                    <input type="radio" name="cp_A03HASUB" value="1"><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_A03HASUB" value="0"><span class="fontTD">N홒</span>&nbsp;
                    <input type="radio" name="cp_A03HASUB" value="0"><span class="fontTD">N홒</span>&nbsp;
       </P></TD>
       <TD><P><span class="aFontePT">A03HASUB<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A03HASUB></P></TD>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
sNot = request.querystring("sF")
sSel = "Select * From CATEGORIAS Where A00IDC1 =" & sNot
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_A01CATEG1               = Trim(objRS("A01CATEG1"))
va_A01CATEG1               = Replace(va_A01CATEG1,"<BR>", chr(13),1)
va_A02EXIBIR               = Trim(objRS("A02EXIBIR"))
va_A03HASUB                = Trim(objRS("A03HASUB"))
va_A03HASUB                = Trim(objRS("A03HASUB"))
va_A03HASUB                = Replace(va_A03HASUB,"<BR>", chr(13),1)
'-----------------------------------------------
%>
                            <table align="center" border=1 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_CATEGORIAS_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">CATEGORIA<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A01CATEG1 %>' name=cp_A01CATEG1></P></TD>
       <TD><P><span class="aFontePT">EXIBIR<br></span>
                    <input type="radio" name="cp_A02EXIBIR" value="1"<% If va_A02EXIBIR = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span><br>
                    <input type="radio" name="cp_A02EXIBIR" value="0"<% If va_A02EXIBIR = "0" Then %> checked<% End If %>><span class="fontTD">N홒</span><br>
       </P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">HA SUB-CATEG<br></span>
                    <input type="radio" name="cp_A03HASUB" value="1"<% If va_A03HASUB = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span><br>
                    <input type="radio" name="cp_A03HASUB" value="0"<% If va_A03HASUB = "0" Then %> checked<% End If %>><span class="fontTD">N홒</span><br>
                    <input type="radio" name="cp_A03HASUB" value="0"<% If va_A03HASUB = "0" Then %> checked<% End If %>><span class="fontTD">N홒</span><br>
       </P></TD>
       <TD><P><span class="aFontePT">A03HASUB<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A03HASUB %>' name=cp_A03HASUB></P></TD>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From CATEGORIAS Where A00IDC1 =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_CATEGORIAS.asp"
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
va_A01CATEG1               = Trim(UploadRequest.Item("cp_A01CATEG1").Item("Value"))
va_A01CATEG1               = Replace(va_A01CATEG1,chr(13),"<BR>" ,1)
va_A02EXIBIR               = Trim(UploadRequest.Item("cp_A02EXIBIR").Item("Value"))
va_A03HASUB                = Trim(UploadRequest.Item("cp_A03HASUB").Item("Value"))
va_A03HASUB                = Trim(UploadRequest.Item("cp_A03HASUB").Item("Value"))
va_A03HASUB                = Replace(va_A03HASUB,chr(13),"<BR>" ,1)
'-----------------------------------------------

If sTipo = "1" Then
sNot = request.querystring("sF")
   sSel = sSel & "INSERT INTO CATEGORIAS(A01CATEG1, A02EXIBIR, A03HASUB, A03HASUB)"
   sSel = sSel & "VALUES ('" & va_A01CATEG1 & "', '" & va_A02EXIBIR & "', '" & va_A03HASUB & "', '" & va_A03HASUB & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CATEGORIAS.asp"

ElseIf sTipo = "2" Then

sNot = request.querystring("sF")
   sSel = sSel & "UPDATE CATEGORIAS SET "
   sSel = sSel & "A01CATEG1 = '" & va_A01CATEG1 & "', A02EXIBIR = '" & va_A02EXIBIR & "', A03HASUB = '" & va_A03HASUB & "', A03HASUB = '" & va_A03HASUB & "' "
   sSel = sSel & "WHERE A00IDC1 = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CATEGORIAS.asp"
Else
   sSel = sSel & "UPDATE CATEGORIAS SET "
   sSel = sSel & "A01CATEG1 = '" & va_A01CATEG1 & "', A02EXIBIR = '" & va_A02EXIBIR & "', A03HASUB = '" & va_A03HASUB & "', A03HASUB = '" & va_A03HASUB & "' "
   sSel = sSel & "WHERE A00IDC1 = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CATEGORIAS.asp"
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

