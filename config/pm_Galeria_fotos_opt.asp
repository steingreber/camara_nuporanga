<%
'<------- /같같같같같같같같같같같같같\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_Galeria_fotos_opt.asp
'CRIADO EM.........: 28/10/2006 12:38:12
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "default.asp"
Dim acao, sNome, sIdc
acao = Request.QueryString("acao")
sIdc = request("id")
If acao = "1" Then
   sNome = "NOVO REGISTRO PARA A GALERIA N? #"& sIdc
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO PARA A GALERIA N? #"& sIdc
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
window.location = "pm_Galeria.asp";
}
function Validator(theForm)
{
if (theForm.cp_a01galeria && !EW_hasValue(theForm.cp_a01galeria, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_a01galeria, "SELECT", "Selecione um valor para o campo GALERIA"))
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
            <table width="605" border="1" cellspacing="0" align="center" bordercolorlight="#808080" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table align="center" border=1 cellspacing=0 width="605" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_Galeria_fotos_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="cp_a01galeria" value="<%= sIdc %>">
						   <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
<!--        <TD><P><span class="aFontePT">GALERIA<br></span>
           <select name="cp_a01galeria" size="1" class="ud_caixa">
<%sLin = "Select * From Galeria"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option selected>Selecione</option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("a00id") %>"><%= objRs("a03festa") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select></td> -->
       <TD><P><span class="aFontePT">FOTO PEQUENA (Tamanho 103x77)<br></span><INPUT class=ud_caixa style="WIDTH:100%" name=cp_a02fotop type="file"></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">FOTO GRANDE (Tamanho 400x300)<br></span><INPUT class=ud_caixa style="WIDTH:100%" name=cp_a03fotog type="file"></P></TD>
</tr>
<tr>
		<td><P><span class="aFontePT">DESCRI플O DA IMAGEM<br></span><input type="text" name="cp_a04descricao" class="ud_caixa" style="WIDTH:100%"></P></td>
</tr>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then %>
<%
On Error Resume Next
sNot = request.querystring("sF")
sSel = "Select * From Galeria_fotos Where a00idg =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a01galeria              = Trim(objRS("a01galeria"))
va_a02fotop                = Trim(objRS("a02fotop"))
va_a03fotog                = Trim(objRS("a03fotog"))
va_a04descricao            = Trim(objRS("a04descricao"))
%>
                            <table width="605" border="1" cellspacing="0" align="center" bordercolorlight="#808080" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
                        <form action="pm_Galeria_fotos_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
	                       <input type="hidden" name="cp_a01galeria" value="<%=va_a01galeria%>">
  					       <INPUT type="hidden" name="hcp_a02fotop" value="<%=va_a02fotop%>">
						   <INPUT type="hidden" name="hcp_a03fotog" value="<%=va_a03fotog%>">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<tr>
       <TD><P><span class="aFontePT">FOTO PEQUENA - <a href="javascript:fotoZoom('<%= "galeria/" & va_a02fotop%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:100%" name=cp_a02fotop type="file"><BR></P>
</TR>
<TR>
       <TD><P><span class="aFontePT">FOTO GRANDE - <a href="javascript:fotoZoom('<%= "galeria/" & va_a03fotog%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:100%" name=cp_a03fotog type="file"><BR></P>
</tr>
<tr>
		<td><P><span class="aFontePT">DESCRI플O DA IMAGEM<br></span><input type="text" name="cp_a04descricao" value="<%= va_a04descricao %>" class="ud_caixa" style="WIDTH:100%"></P></td>
</tr>
                        </form>
                            </table>
<% ElseIf acao = "3" Then

sNot = request.querystring("sF")
sSel = "Delete From Galeria_fotos Where a00idg =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_Galeria.asp"

ElseIf acao = "4" Then
sNot = request.querystring("sF")
Response.Expires = 0
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest RequestBin
Dim sTipo, sCmd
sTipo  = request.querystring("tp")

va_a01galeria              = Trim(UploadRequest.Item("cp_a01galeria").Item("Value"))
va_a02fotop                = Trim(UploadRequest.Item("cp_a02fotop").Item("Value"))
va_a03fotog                = Trim(UploadRequest.Item("cp_a03fotog").Item("Value"))
va_a04descricao            = Trim(UploadRequest.Item("cp_a04descricao").Item("Value"))

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_a02fotop <> "" Then
	Set selectfig = objConect.Execute("SELECT * FROM Galeria_fotos WHERE a02fotop = '" & FileName & "';")
	ContentType0 = UploadRequest.Item("cp_a02fotop").Item("ContentType")
	filepathname0 = UploadRequest.Item("cp_a02fotop").Item("FileName")
	FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
	Set selectfig = objConect.Execute("SELECT * FROM Galeria_fotos WHERE a02fotop = '" & FileName & "';")
	Value0 = UploadRequest.Item("cp_a02fotop").Item("Value")
	Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
	numero0 = instrrev(Request.servervariables("Path_Info"), "/")
	Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("galeria") & "\"&sIdc&"p_" & FileName0)
	For i = 1 To LenB(Value0)
	MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
	MyFile0.Close
	Errou
	va_a02fotop = FileName0
End If

'===============================================
If va_a03fotog <> "" Then
ContentType1 = UploadRequest.Item("cp_a03fotog").Item("ContentType")
filepathname1 = UploadRequest.Item("cp_a03fotog").Item("FileName")
FileName1 = Right(filepathname1, Len(filepathname1) - InStrRev(filepathname1, "\"))
Value1 = UploadRequest.Item("cp_a03fotog").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero1 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile1 = ScriptObject.CreateTextFile(Server.mappath("galeria") & "\"&sIdc&"g_" & FileName1)
For i = 1 To LenB(Value1)
MyFile1.Write Chr(AscB(MidB(Value1, i, 1)))
Next
MyFile1.Close

Errou
va_a03fotog = FileName1
End If

   sSel = sSel & "INSERT INTO Galeria_fotos(a01galeria, a02fotop, a03fotog, a04descricao)"
   sSel = sSel & "VALUES ('" & va_a01galeria & "', '"&sIdc&"p_" & va_a02fotop & "', '"&sIdc&"g_" & va_a03fotog & "', '" & va_a04descricao & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_Galeria.asp"

ElseIf sTipo = "2" Then

'***********************************************
va_a02fotoph                = Trim(UploadRequest.Item("hcp_a02fotop").Item("Value"))
va_a03fotogh                = Trim(UploadRequest.Item("hcp_a03fotog").Item("Value"))
'***********************************************
sNot = request.querystring("sF")
'===============================================
If va_a02fotop <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM Galeria_fotos WHERE a02fotop = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a02fotop").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a02fotop").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM Galeria_fotos WHERE a02fotop = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a02fotop").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("galeria") & "\"&sIdc&"p_" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a02fotop = FileName0
Else
va_a02fotop = va_a02fotoph
End If

'===============================================
If va_a03fotog <> "" Then
ContentType1 = UploadRequest.Item("cp_a03fotog").Item("ContentType")
filepathname1 = UploadRequest.Item("cp_a03fotog").Item("FileName")
FileName1 = Right(filepathname1, Len(filepathname1) - InStrRev(filepathname1, "\"))
Value1 = UploadRequest.Item("cp_a03fotog").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero1 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile1 = ScriptObject.CreateTextFile(Server.mappath("galeria") & "\"&sIdc&"g_" & FileName1)
For i = 1 To LenB(Value1)
MyFile1.Write Chr(AscB(MidB(Value1, i, 1)))
Next
MyFile1.Close

Errou
va_a03fotog = FileName1
Else
va_a03fotog = va_a03fotogh
End If

   sSel = sSel & "UPDATE Galeria_fotos SET "
   sSel = sSel & "a01galeria = '" & va_a01galeria & "', a02fotop = '" & va_a02fotop & "', a03fotog = '" & va_a03fotog & "', a04descricao = '" & va_a04descricao & "' "
   sSel = sSel & "WHERE a00idg = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_Galeria.asp"
Else
   sSel = sSel & "UPDATE Galeria_fotos SET "
   sSel = sSel & "a01galeria = '" & va_a01galeria & "' "
   sSel = sSel & "WHERE a00idg = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_Galeria.asp"
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
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" align="center" cellpadding="0" cellspacing="0" width="571" bordercolordark="black" bordercolorlight="#808080">
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

