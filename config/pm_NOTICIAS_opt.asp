<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_noticias_opt.asp
'CRIADO EM.........: 11/1/2006 16:12:07
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - NOTÍCIAS"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - NOTÍCIAS"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>.::: PageMaster 2.0 :::.</title>
<script language="Javascript1.2" src="editor.js"></script>
<script>_editor_url = "";</script>
<SCRIPT language=Javascript>

function go2url() {
window.location = "pm_noticias.asp";
}
function Validator(theForm)
{
  if (theForm.cp_a01data.value == "")
  {
    alert("O campo DATA não pode ser vazio!");
    theForm.cp_a01data.focus();
    return (false);
  }
  if (theForm.cp_a02titulo.value == "")
  {
    alert("O campo TITULO não pode ser vazio!");
    theForm.cp_a02titulo.focus();
    return (false);
  }
  if (theForm.cp_a03materia.value == "")
  {
    alert("O campo MATERIA não pode ser vazio!");
    theForm.cp_a03materia.focus();
    return (false);
  }
  if (theForm.cp_a04fonte.value == "")
  {
    alert("O campo FONTE não pode ser vazio!");
    theForm.cp_a04fonte.focus();
    return (false);
  }
  if (theForm.cp_a06destaque.value == "")
  {
    alert("O campo DESTAQUE não pode ser vazio!");
    theForm.cp_a06destaque.focus();
    return (false);
  }
  return (true);
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
<body background="images/bg.gif" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000" leftmargin="0" topmargin="0">
<table width="100%" cellpadding="0" cellspacing="0" height="100%">
    <tr>
        <td align="center" valign="middle">
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
            <table align="center" border="1" cellspacing="0" width="605" bordercolordark="white" bordercolordark="#FFFFFF">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table width="605" border="1" cellspacing="0" align="center" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
                        <form action="pm_noticias_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">DATA<br></span><input type="text" name="cp_a01data" value="<%= date %>" class="ud_caixa" style="WIDTH:160px"></P></TD>
       <TD><P><span class="aFontePT">TITULO DA NOTÍCIA<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_a02titulo></P></TD>
</TR>
<TR>
       <td colspan="2">
	   <span class="aFontePT">MATERIA<br></span><TEXTAREA class=ud_caixa style="width:100%; height:250; border-style:none;" name=cp_a03materia></textarea>
	   	<script language="javascript1.2">
		editor_generate('cp_a03materia'); // field, width, height
		</script>	
	   </td>
</TR>
<TR>
       <td><P><span class="aFontePT">FONTE DE ONDE FOI RETIRADA ESTA NOTÍCIA<br></span><INPUT class=ud_caixa style="WIDTH:100%" name=cp_a04fonte></P></td>
       <TD><P><span class="aFontePT">DESTAQUE NA PRIMEIRA PÁGINA<br></span>
           <select name="cp_a06destaque" size="1" class="ud_caixa">
                <option selected>Selecione</option>
                    <option value="S">SIM</option>
                    <option value="N" selected>NÃO</option>
                </select>
       </P></TD>
</TR>
<TR>
       <TD colspan="2"><P><span class="aFontePT">IMAGEM - <font color="#FF0000">(LARGURA MÁXIMA 300 pixel)</font><br></span><INPUT class=ud_caixa style="WIDTH:100%" name=cp_a05img type="file"></P></TD>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then

sNot = request.querystring("sF")
sSel = "Select * From noticias Where a00id =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a01data                 = Trim(objRS("a01data"))
va_a02titulo               = Trim(objRS("a02titulo"))
va_a03materia              = Trim(objRS("a03materia"))
va_a04fonte                = Trim(objRS("a04fonte"))
va_a05img                  = Trim(objRS("a05img"))
va_a06destaque             = Trim(objRS("a06destaque"))
'-----------------------------------------------
%>
                            <table width="605" border="1" cellspacing="0" align="center" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
                        <form action="pm_noticias_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
       <INPUT type="hidden" name="hcp_a05img" value="<%=va_a05img%>">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">DATA<br></span><INPUT class=ud_caixa style="WIDTH:160px" value='<% = va_a01data %>' name=cp_a01data></P></TD>
       <TD><P><span class="aFontePT">TITULO DA NOTÍCIA<br></span><INPUT class=ud_caixa style="WIDTH:250px" value='<% = va_a02titulo %>' name=cp_a02titulo></P></TD>
</TR>
<TR>
       <td colspan="2">
	   <span class="aFontePT">MATERIA<br></span><TEXTAREA class=ud_caixa style="width:100%; height:200; border-style:none;" name=cp_a03materia><% = va_a03materia %></textarea>
	   	<script language="javascript1.2">
		editor_generate('cp_a03materia'); // field, width, height
		</script>
	   </td>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">FONTE DE ONDE FOI RETIRADA ESTA NOTÍCIA<br></span><INPUT class=ud_caixa style="WIDTH:250px" value='<% = va_a04fonte %>' name=cp_a04fonte></P></td>
</TR>
<TR>
       <TD><P><span class="aFontePT">IMAGEM<% If va_a05img<>"" Then %> - <a href="javascript:fotoZoom('<%= "noticias/" & va_a05img%>')">VER AQUIVO</a></span><% End If %><br><INPUT class=ud_caixa style="WIDTH:250px" name=cp_a05img type="file"><BR></P>
       <TD><P><span class="aFontePT">DESTAQUE NA PRIMEIRA PÁGINA<br></span>
           <select name="cp_a06destaque" size="1" class="ud_caixa">
                <option value="<%=va_a06destaque%>" selected><%=va_a06destaque%></option>
                    <option value="S">SIM</option>
                    <option value="N">NÃO</option>
                </select>
       </P></TD>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then

sNot = request.querystring("sF")
sSel = "Delete From noticias Where a00id =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_noticias.asp"

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
va_a01data                 = Trim(UploadRequest.Item("cp_a01data").Item("Value"))
va_a02titulo               = Trim(UploadRequest.Item("cp_a02titulo").Item("Value"))
va_a03materia              = Trim(UploadRequest.Item("cp_a03materia").Item("Value"))
va_a03materia              = Replace(va_a03materia,"'",chr(34),1)
va_a04fonte                = Trim(UploadRequest.Item("cp_a04fonte").Item("Value"))
va_a05img                  = Trim(UploadRequest.Item("cp_a05img").Item("Value"))
va_a06destaque             = Trim(UploadRequest.Item("cp_a06destaque").Item("Value"))
'-----------------------------------------------

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_a05img <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM noticias WHERE a05img = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a05img").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a05img").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM noticias WHERE a05img = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a05img").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("noticias") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a05img = FileName0
End If

   sSel = sSel & "INSERT INTO noticias(a01data, a02titulo, a03materia, a04fonte, a05img, a06destaque)"
   sSel = sSel & "VALUES ('" & va_a01data & "', '" & va_a02titulo & "', '" & va_a03materia & "', '" & va_a04fonte & "', '" & va_a05img & "', '" & va_a06destaque & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_noticias.asp"

ElseIf sTipo = "2" Then

'***********************************************
va_a05imgh                  = Trim(UploadRequest.Item("hcp_a05img").Item("Value"))
va_a05imgh                  = Replace(va_a05imgh,chr(13),"<BR>" ,1)
'***********************************************
sNot = request.querystring("sF")
'===============================================
If va_a05img <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM noticias WHERE a05img = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a05img").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a05img").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM noticias WHERE a05img = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a05img").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("noticias") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a05img = FileName0
Else
va_a05img = va_a05imgh
End If

   sSel = sSel & "UPDATE noticias SET "
   sSel = sSel & "a01data = '" & va_a01data & "', a02titulo = '" & va_a02titulo & "', a03materia = '" & va_a03materia & "', a04fonte = '" & va_a04fonte & "', a05img = '" & va_a05img & "', a06destaque = '" & va_a06destaque & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_noticias.asp"
Else
   sSel = sSel & "UPDATE noticias SET "
   sSel = sSel & "a01data = '" & va_a01data & "', a02titulo = '" & va_a02titulo & "', a03materia = '" & va_a03materia & "', a04fonte = '" & va_a04fonte & "', a06destaque = '" & va_a06destaque & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_noticias.asp"
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
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" align="center" cellpadding="0" cellspacing="0" width="571" bordercolordark="black" bordercolordark="#FFFFFF">
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

