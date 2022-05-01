<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_Galeria_opt.asp
'CRIADO EM.........: 28/10/2006 18:51:22
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "default.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - GALERIA DE FOTOS DO MUNICÍPIO"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - GALERIA DE FOTOS DO MUNICÍPIO"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>.::: PageMaster 2.0 :::.</title>
<meta name="robots" content="none"/>
<script language="JavaScript" src="lu.js"></script>
<SCRIPT language=Javascript>

function go2url() {
window.location = "pm_Galeria.asp";
}
function Validator(theForm)
{
if (theForm.cp_a03festa && !EW_hasValue(theForm.cp_a03festa, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a03festa, "TEXT", "Informe um valor para o campo NOME DA FESTA"))
            return false;
    }
if (theForm.cp_a01local && !EW_hasValue(theForm.cp_a01local, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a01local, "TEXT", "Informe um valor para o campo LOCAL"))
            return false;
    }
  if (theForm.cp_a02data && !EW_hasValue(theForm.cp_a02data, "TEXT" )) {
          if (!EW_onError(theForm, theForm.cp_a02data, "TEXT", "Informe a data correta para o campo DATA"))
              return false;
      }
  if (theForm.cp_a02data && !EW_checkeurodate(theForm.cp_a02data.value)) {
      if (!EW_onError(theForm, theForm.cp_a02data, "TEXT", "Data inválida para o campo DATA"))
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
            <table align="center" border="1" cellspacing="0" width="605" bordercolordark="white" bordercolorlight="#808080">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table width="605" border="1" cellspacing="0" align="center" bordercolorlight="#808080" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
                        <form action="pm_Galeria_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">NOME GALERIA<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a03festa></P></TD>
       <TD><P><span class="aFontePT">LOCAL<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a01local></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">DATA<br></span><input type="text" name="cp_a02data" value="<%=date%>" class="ud_caixa" style="WIDTH:150px">&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" onClick="popUpCalendar(this, this.form.cp_a02data,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">FOTO<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a04foto type="file"></P></TD>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then %>
<%
On Error Resume Next
sNot = request.querystring("sF")
sSel = "Select * From Galeria Where a00id =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a03festa                = Trim(objRS("a03festa"))
va_a01local                = Trim(objRS("a01local"))
va_a02data                 = Trim(objRS("a02data"))
va_a04foto                 = Trim(objRS("a04foto"))
%>
                            <table width="605" border="1" cellspacing="0" align="center" bordercolorlight="#808080" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
                        <form action="pm_Galeria_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">NOME GALERIA<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a03festa %>' name=cp_a03festa></P></TD>
       <TD><P><span class="aFontePT">LOCAL<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a01local %>' name=cp_a01local></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">DATA<br></span><INPUT class=ud_caixa style="WIDTH:150px" value='<% = va_a02data %>' name=cp_a02data>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" onClick="popUpCalendar(this, this.form.cp_a02data,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">FOTO - <a href="javascript:fotoZoom('<%= "galeria/" & va_a04foto%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a04foto type="file"><BR></P>
       <INPUT type="hidden" name="hcp_a04foto" value="<%=va_a04foto%>">
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then %>
<%
sNot = request.querystring("sF")
sSel = "Delete From Galeria Where a00id =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_Galeria.asp"
%>
<%
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
va_a03festa                = Trim(UploadRequest.Item("cp_a03festa").Item("Value"))
va_a01local                = Trim(UploadRequest.Item("cp_a01local").Item("Value"))
va_a02data                 = Trim(UploadRequest.Item("cp_a02data").Item("Value"))
va_a04foto                 = Trim(UploadRequest.Item("cp_a04foto").Item("Value"))

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_a04foto <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM Galeria WHERE a04foto = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a04foto").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a04foto").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM Galeria WHERE a04foto = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a04foto").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("galeria") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a04foto = FileName0
End If

   sSel = sSel & "INSERT INTO Galeria(a03festa, a01local, a02data, a04foto)"
   sSel = sSel & "VALUES ('" & va_a03festa & "', '" & va_a01local & "', '" & va_a02data & "', '" & va_a04foto & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_Galeria.asp"

ElseIf sTipo = "2" Then

'***********************************************
va_a04fotoh                 = Trim(UploadRequest.Item("hcp_a04foto").Item("Value"))
va_a04fotoh                 = Replace(va_a04fotoh,chr(13),"<BR>" ,1)
'***********************************************
sNot = request.querystring("sF")
'===============================================
If va_a04foto <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM Galeria WHERE a04foto = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a04foto").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a04foto").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM Galeria WHERE a04foto = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a04foto").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("galeria") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a04foto = FileName0
Else
va_a04foto = va_a04fotoh
End If

   sSel = sSel & "UPDATE Galeria SET "
   sSel = sSel & "a03festa = '" & va_a03festa & "', a01local = '" & va_a01local & "', a02data = '" & va_a02data & "', a04foto = '" & va_a04foto & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_Galeria.asp"
Else
   sSel = sSel & "UPDATE Galeria SET "
   sSel = sSel & "a03festa = '" & va_a03festa & "', a01local = '" & va_a01local & "', a02data = '" & va_a02data & "' "
   sSel = sSel & "WHERE a00id = " & sNot
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
            <table width="571" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#808080" bordercolordark="#000000" bgcolor="#FFFFFF" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid">
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

