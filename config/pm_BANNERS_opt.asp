<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_banners_opt.asp
'CRIADO EM.........: 26/11/2006 15:41:08
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - BANNERS"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - BANNERS"
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
window.location = "pm_banners.asp";
}
function Validator(theForm)
{
  if (theForm.cp_a04data && !EW_hasValue(theForm.cp_a04data, "TEXT" )) {
          if (!EW_onError(theForm, theForm.cp_a04data, "TEXT", "Informe a data correta para o campo DATA"))
              return false;
      }
  if (theForm.cp_a04data && !EW_checkeurodate(theForm.cp_a04data.value)) {
      if (!EW_onError(theForm, theForm.cp_a04data, "TEXT", "Data inválida para o campo DATA"))
          return false;
      }
if (theForm.cp_a01descr && !EW_hasValue(theForm.cp_a01descr, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a01descr, "TEXT", "Informe um valor para o campo DESCRIÇÃO"))
            return false;
    }
if (theForm.cp_a02link && !EW_hasValue(theForm.cp_a02link, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a02link, "TEXT", "Informe um valor para o campo LINK"))
            return false;
    }
  if (theForm.cp_a06abrir && !EW_hasValue(theForm.cp_a06abrir, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_a06abrir, "RADIO", "Selecione um valor para ABRIR EM!"))
            return false;
    }
  if (theForm.cp_a05exibir && !EW_hasValue(theForm.cp_a05exibir, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_a05exibir, "RADIO", "Selecione um valor para EXIBIR?!"))
            return false;
    }
  if (theForm.cp_a07tamanho && !EW_hasValue(theForm.cp_a07tamanho, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_a07tamanho, "RADIO", "Selecione o tamanho do BANNER!"))
            return false;
    }
  if (theForm.cp_a08tipo && !EW_hasValue(theForm.cp_a08tipo, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_a08tipo, "RADIO", "Selecione o TIPO do BANNER!"))
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
                            <table align="center" border=1 cellspacing=0 width="605" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_banners_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<tr>
       <TD><P><span class="aFontePT">DATA<br></span><input type="text" name="cp_a04data" value="<%= date %>" class="ud_caixa" style="WIDTH:150px">&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" onClick="popUpCalendar(this, this.form.cp_a04data,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">DESCRIÇÃO<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a01descr></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">LINK<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a02link></P></TD>
       <TD><P><span class="aFontePT">ABRIR EM<br></span>
                    <input type="radio" name="cp_a06abrir" value="1" checked><span class="fontTD">MESMA PÁGINA</span>&nbsp;
                    <input type="radio" name="cp_a06abrir" value="0"><span class="fontTD">NOVA PÁGINA</span>
       </P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">EXIBIR ESTE BANNER?<br></span>
                    <input type="radio" name="cp_a05exibir" value="1" checked><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_a05exibir" value="0"><span class="fontTD">NÃO</span>
       </P></TD>
       <TD><P><span class="aFontePT">IMAGEM (TAMANHO 400X60 ou 130x54)<br></span><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a03imagen type="file"></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">TAMANHO DO BANNER EM pixel<br></span>
                    <input type="radio" name="cp_a07tamanho" value="1"><span class="fontTD">400X60 Top</span>&nbsp;
                    <input type="radio" name="cp_a07tamanho" value="0"><span class="fontTD">130x54 Lateral</span>&nbsp;
					<input type="radio" name="cp_a07tamanho" value="2"><span class="fontTD">Centro</span>&nbsp;
       </P></TD>
       <TD><P><span class="aFontePT">TIPO DE BANNER<br></span>
                    <input type="radio" name="cp_a08tipo" value="1"><span class="fontTD">FLASH</span>&nbsp;
                    <input type="radio" name="cp_a08tipo" value="0"><span class="fontTD">JPEG</span>
       </P></TD>
</TR>
<tr>
		<td colspan="2">
		<P><span class="aFontePT">DIMENSÕES (SOMENTE PARA FORMATO FLASH E TAMANHO CENTRO)<br>
		LARGURA:<INPUT class=ud_caixa style="WIDTH:90px" name=cp_a10largura>&nbsp;
		ALTURA:<INPUT class=ud_caixa style="WIDTH:90px" name=cp_a09altura></span>
		</P>
		</td>
</tr>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then %>
<%
On Error Resume Next
sNot = request.querystring("sF")
sSel = "Select * From banners Where a00id =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a04data                 = Trim(objRS("a04data"))
va_a01descr                = Trim(objRS("a01descr"))
va_a02link                 = Trim(objRS("a02link"))
va_a06abrir                = Trim(objRS("a06abrir"))
va_a05exibir               = Trim(objRS("a05exibir"))
va_a03imagen               = Trim(objRS("a03imagen"))
va_a07tamanho              = Trim(objRS("a07tamanho"))
va_a08tipo                 = Trim(objRS("a08tipo"))
va_a09altura               = Trim(objRS("a09altura"))
va_a10largura              = Trim(objRS("a10largura"))

%>
                            <table align="center" border=1 cellspacing=0 width="605" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_banners_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                                      <INPUT type="hidden" name="hcp_a03imagen" value="<%=va_a03imagen%>">
							   <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">DATA<br></span><INPUT class=ud_caixa style="WIDTH:150px" value='<% = va_a04data %>' name=cp_a04data>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" onClick="popUpCalendar(this, this.form.cp_a04data,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">DESCRIÇÃO<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a01descr %>' name=cp_a01descr></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">LINK<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a02link %>' name=cp_a02link></P></TD>
       <TD><P><span class="aFontePT">ABRIR EM<br></span>
                    <input type="radio" name="cp_a06abrir" value="1"<% If va_a06abrir = "1" Then %> checked<% End If %>><span class="fontTD">MESMA PÁGINA</span>&nbsp;
                    <input type="radio" name="cp_a06abrir" value="0"<% If va_a06abrir = "0" Then %> checked<% End If %>><span class="fontTD">NOVA PÁGINA</span>
       </P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">EXIBIR ESTE BANNER?<br></span>
                    <input type="radio" name="cp_a05exibir" value="1"<% If va_a05exibir = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_a05exibir" value="0"<% If va_a05exibir = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>
       </P></TD>
       <TD><P><span class="aFontePT">IMAGEM - <a href="javascript:fotoZoom('<%= "banners/" & va_a03imagen%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:290px" name=cp_a03imagen type="file"><BR></P>
</TR>
<tr>
       <TD><P><span class="aFontePT">TAMANHO DO BANNER EM pixel<br></span>
                    <input type="radio" name="cp_a07tamanho" value="1"<% If va_a07tamanho= "1" Then %> checked<% End If %>><span class="fontTD">400x60 Top</span>&nbsp;
                    <input type="radio" name="cp_a07tamanho" value="0"<% If va_a07tamanho = "0" Then %> checked<% End If %>><span class="fontTD">130x54 Lateral</span>&nbsp;
                    <input type="radio" name="cp_a07tamanho" value="2"<% If va_a07tamanho = "2" Then %> checked<% End If %>><span class="fontTD">Centro</span>
       </P></TD>
       <TD><P><span class="aFontePT">TIPO DE BANNER?<br></span>
                    <input type="radio" name="cp_a08tipo" value="1"<% If va_a08tipo = "1" Then %> checked<% End If %>><span class="fontTD">FLASH</span>&nbsp;
                    <input type="radio" name="cp_a08tipo" value="0"<% If va_a08tipo = "0" Then %> checked<% End If %>><span class="fontTD">JPEG</span>
       </P></TD>
</tr>
<tr>
		<td colspan="2">
		<P><span class="aFontePT">DIMENSÕES (SOMENTE PARA FORMATO FLASH E TAMANHO CENTRO)<br>
		LARGURA:<INPUT class=ud_caixa style="WIDTH:90px" value='<% = va_a10largura %>' name=cp_a10largura>&nbsp;
		ALTURA:<INPUT class=ud_caixa style="WIDTH:90px" value='<% = va_a09altura %>' name=cp_a09altura></span>
		</P>
		</td>
</tr>


                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From banners Where a00id =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_banners.asp"

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
va_a04data                 = Trim(UploadRequest.Item("cp_a04data").Item("Value"))
va_a01descr                = Trim(UploadRequest.Item("cp_a01descr").Item("Value"))
va_a02link                 = Trim(UploadRequest.Item("cp_a02link").Item("Value"))
va_a06abrir                = Trim(UploadRequest.Item("cp_a06abrir").Item("Value"))
va_a05exibir               = Trim(UploadRequest.Item("cp_a05exibir").Item("Value"))
va_a03imagen               = Trim(UploadRequest.Item("cp_a03imagen").Item("Value"))
va_a07tamanho              = Trim(UploadRequest.Item("cp_a07tamanho").Item("Value"))
va_a08tipo                 = Trim(UploadRequest.Item("cp_a08tipo").Item("Value"))
va_a10largura              = Trim(UploadRequest.Item("cp_a10largura").Item("Value"))
va_a09altura               = Trim(UploadRequest.Item("cp_a09altura").Item("Value"))

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_a03imagen <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM banners WHERE a03imagen = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a03imagen").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a03imagen").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM banners WHERE a03imagen = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a03imagen").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("banners") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a03imagen = FileName0
End If

   sSel = sSel & "INSERT INTO banners(a04data, a01descr, a02link, a06abrir, a05exibir, a03imagen, a07tamanho, a08tipo, a09altura, a10largura)"
   sSel = sSel & "VALUES ('" & va_a04data & "', '" & va_a01descr & "', '" & va_a02link & "', '" & va_a06abrir & "', '" & va_a05exibir & "', '" & va_a03imagen & "', '" & va_a07tamanho & "', '" & va_a08tipo & "', '" & va_a09altura & "', '" & va_a10largura & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_banners.asp"

ElseIf sTipo = "2" Then

'***********************************************
va_a03imagenh               = Trim(UploadRequest.Item("hcp_a03imagen").Item("Value"))
va_a03imagenh               = Replace(va_a03imagenh,chr(13),"<BR>" ,1)
'***********************************************
sNot = request.querystring("sF")
'===============================================
If va_a03imagen <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM banners WHERE a03imagen = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a03imagen").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a03imagen").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM banners WHERE a03imagen = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a03imagen").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("banners") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a03imagen = FileName0
Else
va_a03imagen = va_a03imagenh
End If

   sSel = sSel & "UPDATE banners SET "
   sSel = sSel & "a04data = '" & va_a04data & "', a01descr = '" & va_a01descr & "', a02link = '" & va_a02link & "', a06abrir = '" & va_a06abrir & "', a05exibir = '" & va_a05exibir & "', a03imagen = '" & va_a03imagen & "', a07tamanho = '" & va_a07tamanho & "', a08tipo = '" & va_a08tipo & "', a09altura = '" & va_a09altura & "', a10largura = '" & va_a10largura & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_banners.asp"
Else
   sSel = sSel & "UPDATE banners SET "
   sSel = sSel & "a04data = '" & va_a04data & "', a01descr = '" & va_a01descr & "', a02link = '" & va_a02link & "', a06abrir = '" & va_a06abrir & "', a05exibir = '" & va_a05exibir & "', a07tamanho = '" & va_a07tamanho & "', a08tipo = '" & va_a08tipo & "', a09altura = '" & va_a09altura & "', a10largura = '" & va_a10largura & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_banners.asp"
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

