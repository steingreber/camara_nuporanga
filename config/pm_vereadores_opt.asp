<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_vereadores_opt.asp
'CRIADO EM.........: 17/11/2005 10:48:06
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - VEREADORES"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - VEREADORES"
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
window.location = "pm_vereadores.asp";
}
function Validator(theForm)
{
  if (theForm.cp_a01nome.value == "")
  {
    alert("O campo NOME não pode ser vazio!");
    theForm.cp_a01nome.focus();
    return (false);
  }
  if (theForm.cp_a02email.value == "")
  {
    alert("O campo EMAIL não pode ser vazio!");
    theForm.cp_a02email.focus();
    return (false);
  }
  if (theForm.cp_a03partido.value == "")
  {
    alert("O campo PARTIDO não pode ser vazio!");
    theForm.cp_a03partido.focus();
    return (false);
  }
  if (theForm.cp_a04descricao.value == "")
  {
    alert("O campo DESCRICAO não pode ser vazio!");
    theForm.cp_a04descricao.focus();
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
<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000" leftmargin=0 topmargin=0>
<table width="100%" cellpadding="0" cellspacing="0" height="100%">
    <tr>
        <td align="center" valign="middle">
<table cellpadding="0" cellspacing="0" width="610" align="center">
    <tr>
        <td class="tabelatitulo" align="center" height="50">
            <p><span class="aFonte">-:- <%= sNome %> -:-</span></p>
        </td>
    </tr>
    <tr>
        <td height="5"" class="tabelaitem">
        </td>
    </tr>
    <tr>
        <td" align="center" valign="top">
            <table width="605" border="1" cellspacing="0" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table width="605" border="1" cellspacing="0" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
                        <form action="pm_vereadores_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                               </tr>
<tr>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a01nome></P></TD>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a02email></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">PARTIDO<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a03partido></P></TD>
       <TD><P><span class="aFontePT">FOTO<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a05foto type="file"></P></TD>
</TR>
<TR>
       <td colspan="2">
	   <span class="aFontePT">DESCRICAO<br></span><TEXTAREA class=ud_caixa style="width:100%; height:100; border-style:none;" name=cp_a04descricao></textarea>
	   	<script language="javascript1.2">
		editor_generate('cp_a04descricao'); // field, width, height
		</script>
	   </td>
</tr>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
sNot = request.querystring("sF")
sSel = "Select * From vereadores Where a00id =" & sNot
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a01nome                 = Trim(objRS("a01nome"))
va_a02email                = Trim(objRS("a02email"))
va_a03partido              = Trim(objRS("a03partido"))
va_a05foto                 = Trim(objRS("a05foto"))
va_a04descricao            = Trim(objRS("a04descricao"))
%>
                            <table width="605" border="1" cellspacing="0" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
                        <form action="pm_vereadores_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
       <INPUT type="hidden" name="hcp_a05foto" value="<%=va_a05foto%>">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a01nome %>' name=cp_a01nome></P></TD>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a02email %>' name=cp_a02email></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">PARTIDO<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a03partido %>' name=cp_a03partido></P></TD>
       <TD><P><span class="aFontePT">FOTO - <a href="javascript:fotoZoom('<%= "vereadores/" & va_a05foto%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a05foto type="file"><BR></P>
</TR>
<TR>
       <td colspan="2">
	   <span class="aFontePT">DESCRICAO<br></span><TEXTAREA class=ud_caixa style="width:100%; height:100; border-style:none;" name=cp_a04descricao><% = va_a04descricao %></textarea>
	   	<script language="javascript1.2">
		editor_generate('cp_a04descricao'); // field, width, height
		</script>
	   </td>
</TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From vereadores Where a00id =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_vereadores.asp"

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
va_a01nome                 = Trim(UploadRequest.Item("cp_a01nome").Item("Value"))
va_a02email                = Trim(UploadRequest.Item("cp_a02email").Item("Value"))
va_a03partido              = Trim(UploadRequest.Item("cp_a03partido").Item("Value"))
va_a05foto                 = Trim(UploadRequest.Item("cp_a05foto").Item("Value"))
va_a04descricao            = Trim(UploadRequest.Item("cp_a04descricao").Item("Value"))

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_a05foto <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM vereadores WHERE a05foto = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a05foto").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a05foto").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM vereadores WHERE a05foto = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a05foto").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("vereadores") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a05foto = FileName0
End If

   sSel = sSel & "INSERT INTO vereadores(a01nome, a02email, a03partido, a05foto, a04descricao)"
   sSel = sSel & "VALUES ('" & va_a01nome & "', '" & va_a02email & "', '" & va_a03partido & "', '" & va_a05foto & "', '" & va_a04descricao & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_vereadores.asp"

ElseIf sTipo = "2" Then

'***********************************************
va_a05fotoh                 = Trim(UploadRequest.Item("hcp_a05foto").Item("Value"))
va_a05fotoh                 = Replace(va_a05fotoh,chr(13),"<BR>" ,1)
'***********************************************
sNot = request.querystring("sF")
'===============================================
If va_a05foto <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM vereadores WHERE a05foto = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a05foto").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a05foto").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM vereadores WHERE a05foto = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a05foto").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("vereadores") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a05foto = FileName0
Else
va_a05foto = va_a05fotoh
End If

   sSel = sSel & "UPDATE vereadores SET "
   sSel = sSel & "a01nome = '" & va_a01nome & "', a02email = '" & va_a02email & "', a03partido = '" & va_a03partido & "', a05foto = '" & va_a05foto & "', a04descricao = '" & va_a04descricao & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_vereadores.asp"
Else
   sSel = sSel & "UPDATE vereadores SET "
   sSel = sSel & "a01nome = '" & va_a01nome & "', a02email = '" & va_a02email & "', a03partido = '" & va_a03partido & "', a04descricao = '" & va_a04descricao & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_vereadores.asp"
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
            <table width="571" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#C0C0C0" bordercolordark="#000000" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid">
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

