<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_turismo_opt.asp
'CRIADO EM.........: 31/3/2006 15:39:22
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - TURISMO"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - TURISMO"
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
window.location = "pm_turismo.asp";
}
function Validator(theForm)
{
  if (theForm.cp_A01CATEG.value == "")
  {
    alert("O campo CATEGORIA não pode ser vazio!");
    theForm.cp_A01CATEG.focus();
    return (false);
  }
  if (theForm.cp_A02NOME.value == "")
  {
    alert("O campo NOME não pode ser vazio!");
    theForm.cp_A02NOME.focus();
    return (false);
  }
  if (theForm.cp_A03END.value == "")
  {
    alert("O campo ENDEREÇO não pode ser vazio!");
    theForm.cp_A03END.focus();
    return (false);
  }
  if (theForm.cp_A04FONE.value == "")
  {
    alert("O campo FONE não pode ser vazio!");
    theForm.cp_A04FONE.focus();
    return (false);
  }
  if (theForm.cp_A05EMAIL.value == "")
  {
    alert("O campo EMAIL não pode ser vazio!");
    theForm.cp_A05EMAIL.focus();
    return (false);
  }
  if (theForm.cp_A06WEB.value == "")
  {
    alert("O campo WEB não pode ser vazio!");
    theForm.cp_A06WEB.focus();
    return (false);
  }
  if (theForm.cp_A07DESC.value == "")
  {
    alert("O campo DESCRIÇÃO não pode ser vazio!");
    theForm.cp_A07DESC.focus();
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
                        <form action="pm_turismo_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">CATEGORIA<br></span>
           <select name="cp_A01CATEG" size="1" class="ud_caixa">
<%sLin = "Select * From turismo_categ"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option selected>Selecione</option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("A01IDTC") %>" selected><%= objRs("A01CATEG") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>&nbsp;<a href="pm_turismo_categ_opt.asp?acao=1"><img src="images/add-comment-red.gif" alt="NOVA CATEGORIA" border="0" align="absmiddle"></a>
       </P></TD>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A02NOME></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">ENDEREÇO<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A03END></P></TD>
       <TD><P><span class="aFontePT">FONE<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A04FONE></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A05EMAIL></P></TD>
       <TD><P><span class="aFontePT">PÁGINA WEB<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A06WEB></P></TD>
</TR>
<TR>
       <td colspan="2">
	   <span class="aFontePT">DESCRIÇÃO<br></span><TEXTAREA class=ud_caixa style="width:100%; height:100; border-style:none;" name=cp_A07DESC></textarea>
	   	<script language="javascript1.2">
		editor_generate('cp_A07DESC'); // field, width, height
		</script>
	   </td>
</tr>
<TR>
	   <TD><P><span class="aFontePT">IMGAGEM 01<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A08IMG01 type="file"></P></TD>
       <TD><P><span class="aFontePT">IMGAGEM 02<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A09IMG02 type="file"></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">IMGAGEM 03<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A10IMG03 type="file"></P></TD>
       <TD><P><span class="aFontePT">IMGAGEM 04<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A11IMG04 type="file"></P></TD>
</tr>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
On Error Resume Next
sNot = request.querystring("sF")
sSel = "Select * From turismo Where A00IDTUR =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_A01CATEG                = Trim(objRS("A01CATEG"))
va_A02NOME                 = Trim(objRS("A02NOME"))
va_A03END                  = Trim(objRS("A03END"))
va_A04FONE                 = Trim(objRS("A04FONE"))
va_A05EMAIL                = Trim(objRS("A05EMAIL"))
va_A06WEB                  = Trim(objRS("A06WEB"))
va_A07DESC                 = Trim(objRS("A07DESC"))
va_A08IMG01                = Trim(objRS("A08IMG01"))
va_A09IMG02                = Trim(objRS("A09IMG02"))
va_A10IMG03                = Trim(objRS("A10IMG03"))
va_A11IMG04                = Trim(objRS("A11IMG04"))
%>
                            <table align="center" border=1 cellspacing=0 width="605" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_turismo_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
       <INPUT type="hidden" name="hcp_A08IMG01" value="<%=va_A08IMG01%>">
       <INPUT type="hidden" name="hcp_A10IMG03" value="<%=va_A10IMG03%>">
       <INPUT type="hidden" name="hcp_A09IMG02" value="<%=va_A09IMG02%>">
       <INPUT type="hidden" name="hcp_A11IMG04" value="<%=va_A11IMG04%>">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">CATEGORIA<br></span>
           <select name="cp_A01CATEG" size="1" class="ud_caixa">
<%
sLin = "Select * From turismo_categ"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("A01IDTC") %>"<% If Trim(objRs(0)) = Trim(va_A01CATEG) Then%> selected<% End If %>><%= objRs("A01IDTC") & "-" & objRs("A01CATEG") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>&nbsp;<a href="pm_turismo_categ_opt.asp?acao=1"><img src="images/add-comment-red.gif" alt="NOVA CATEGORIA" border="0" align="absmiddle"></a>
       </P></TD>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT class=ud_caixa style="WIDTH:250px" value='<% = va_A02NOME %>' name=cp_A02NOME></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">ENDEREÇO<br></span><INPUT class=ud_caixa style="WIDTH:250px" value='<% = va_A03END %>' name=cp_A03END></P></TD>
       <TD><P><span class="aFontePT">FONE<br></span><INPUT class=ud_caixa style="WIDTH:250px" value='<% = va_A04FONE %>' name=cp_A04FONE></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT class=ud_caixa style="WIDTH:250px" value='<% = va_A05EMAIL %>' name=cp_A05EMAIL></P></TD>
       <TD><P><span class="aFontePT">PÁGINA WEB<br></span><INPUT class=ud_caixa style="WIDTH:250px" value='<% = va_A06WEB %>' name=cp_A06WEB></P></TD>
</TR>
<TR>
       <td colspan="2">
	   <span class="aFontePT">DESCRIÇÃO<br></span><TEXTAREA class=ud_caixa style="width:100%; height:100; border-style:none;" name=cp_A07DESC><% = va_A07DESC %></textarea>
	   	<script language="javascript1.2">
		editor_generate('cp_A07DESC'); // field, width, height
		</script>
	   </td>
</TR>
	   <TD><P><span class="aFontePT">IMGAGEM 01 - <a href="javascript:fotoZoom('<%= "turismo/" & va_A08IMG01%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A08IMG01 type="file"><BR></P>
       <TD><P><span class="aFontePT">IMGAGEM 02 - <a href="javascript:fotoZoom('<%= "turismo/" & va_A09IMG02%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A09IMG02 type="file"><BR></P>
</TR>
<TR>
       <TD><P><span class="aFontePT">IMGAGEM 03 - <a href="javascript:fotoZoom('<%= "turismo/" & va_A10IMG03%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A10IMG03 type="file"><BR></P>
	   <TD><P><span class="aFontePT">IMGAGEM 04 - <a href="javascript:fotoZoom('<%= "turismo/" & va_A11IMG04%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:250px" name=cp_A11IMG04 type="file"><BR></P>
</TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From turismo Where A00IDTUR =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_turismo.asp"

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
va_A01CATEG                = Trim(UploadRequest.Item("cp_A01CATEG").Item("Value"))
va_A02NOME                 = Trim(UploadRequest.Item("cp_A02NOME").Item("Value"))
va_A03END                  = Trim(UploadRequest.Item("cp_A03END").Item("Value"))
va_A04FONE                 = Trim(UploadRequest.Item("cp_A04FONE").Item("Value"))
va_A05EMAIL                = Trim(UploadRequest.Item("cp_A05EMAIL").Item("Value"))
va_A06WEB                  = Trim(UploadRequest.Item("cp_A06WEB").Item("Value"))
va_A07DESC                 = Trim(UploadRequest.Item("cp_A07DESC").Item("Value"))
va_A08IMG01                = Trim(UploadRequest.Item("cp_A08IMG01").Item("Value"))
va_A09IMG02                = Trim(UploadRequest.Item("cp_A09IMG02").Item("Value"))
va_A10IMG03                = Trim(UploadRequest.Item("cp_A10IMG03").Item("Value"))
va_A11IMG04                = Trim(UploadRequest.Item("cp_A11IMG04").Item("Value"))

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_A08IMG01 <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM turismo WHERE A08IMG01 = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_A08IMG01").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_A08IMG01").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM turismo WHERE A08IMG01 = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_A08IMG01").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("turismo") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_A08IMG01 = FileName0
End If

'===============================================
If va_A09IMG02 <> "" Then
ContentType1 = UploadRequest.Item("cp_A09IMG02").Item("ContentType")
filepathname1 = UploadRequest.Item("cp_A09IMG02").Item("FileName")
FileName1 = Right(filepathname1, Len(filepathname1) - InStrRev(filepathname1, "\"))
Value1 = UploadRequest.Item("cp_A09IMG02").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero1 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile1 = ScriptObject.CreateTextFile(Server.mappath("turismo") & "\" & FileName1)
For i = 1 To LenB(Value1)
MyFile1.Write Chr(AscB(MidB(Value1, i, 1)))
Next
MyFile1.Close

Errou
va_A09IMG02 = FileName1
End If

'===============================================
If va_A10IMG03 <> "" Then
ContentType2 = UploadRequest.Item("cp_A10IMG03").Item("ContentType")
filepathname2 = UploadRequest.Item("cp_A10IMG03").Item("FileName")
FileName2 = Right(filepathname2, Len(filepathname2) - InStrRev(filepathname2, "\"))
Value2 = UploadRequest.Item("cp_A10IMG03").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero2 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile2 = ScriptObject.CreateTextFile(Server.mappath("turismo") & "\" & FileName2)
For i = 1 To LenB(Value2)
MyFile2.Write Chr(AscB(MidB(Value2, i, 1)))
Next
MyFile2.Close

Errou
va_A10IMG03 = FileName2
End If

'===============================================
If va_A11IMG04 <> "" Then
ContentType3 = UploadRequest.Item("cp_A11IMG04").Item("ContentType")
filepathname3 = UploadRequest.Item("cp_A11IMG04").Item("FileName")
FileName3 = Right(filepathname3, Len(filepathname3) - InStrRev(filepathname3, "\"))
Value3 = UploadRequest.Item("cp_A11IMG04").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero3 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile3 = ScriptObject.CreateTextFile(Server.mappath("turismo") & "\" & FileName3)
For i = 1 To LenB(Value3)
MyFile3.Write Chr(AscB(MidB(Value3, i, 1)))
Next
MyFile3.Close

Errou
va_A11IMG04 = FileName3
End If

   sSel = sSel & "INSERT INTO turismo(A01CATEG, A02NOME, A03END, A04FONE, A05EMAIL, A06WEB, A07DESC, A08IMG01, A09IMG02, A10IMG03, A11IMG04)"
   sSel = sSel & "VALUES ('" & va_A01CATEG & "', '" & va_A02NOME & "', '" & va_A03END & "', '" & va_A04FONE & "', '" & va_A05EMAIL & "', '" & va_A06WEB & "', '" & va_A07DESC & "', '" & va_A08IMG01 & "', '" & va_A09IMG02 & "', '" & va_A10IMG03 & "', '" & va_A11IMG04 & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_turismo.asp"

ElseIf sTipo = "2" Then

'***********************************************
va_A08IMG01h                = Trim(UploadRequest.Item("hcp_A08IMG01").Item("Value"))
va_A08IMG01h                = Replace(va_A08IMG01h,chr(13),"<BR>" ,1)
'***********************************************
va_A09IMG02h                = Trim(UploadRequest.Item("hcp_A09IMG02").Item("Value"))
va_A09IMG02h                = Replace(va_A09IMG02h,chr(13),"<BR>" ,1)
'***********************************************
va_A10IMG03h                = Trim(UploadRequest.Item("hcp_A10IMG03").Item("Value"))
va_A10IMG03h                = Replace(va_A10IMG03h,chr(13),"<BR>" ,1)
'***********************************************
va_A11IMG04h                = Trim(UploadRequest.Item("hcp_A11IMG04").Item("Value"))
va_A11IMG04h                = Replace(va_A11IMG04h,chr(13),"<BR>" ,1)
'***********************************************
sNot = request.querystring("sF")
'===============================================
If va_A08IMG01 <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM turismo WHERE A08IMG01 = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_A08IMG01").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_A08IMG01").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM turismo WHERE A08IMG01 = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_A08IMG01").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("turismo") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_A08IMG01 = FileName0
Else
va_A08IMG01 = va_A08IMG01h
End If

'===============================================
If va_A09IMG02 <> "" Then
ContentType1 = UploadRequest.Item("cp_A09IMG02").Item("ContentType")
filepathname1 = UploadRequest.Item("cp_A09IMG02").Item("FileName")
FileName1 = Right(filepathname1, Len(filepathname1) - InStrRev(filepathname1, "\"))
Value1 = UploadRequest.Item("cp_A09IMG02").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero1 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile1 = ScriptObject.CreateTextFile(Server.mappath("turismo") & "\" & FileName1)
For i = 1 To LenB(Value1)
MyFile1.Write Chr(AscB(MidB(Value1, i, 1)))
Next
MyFile1.Close

Errou
va_A09IMG02 = FileName1
Else
va_A09IMG02 = va_A09IMG02h
End If

'===============================================
If va_A10IMG03 <> "" Then
ContentType2 = UploadRequest.Item("cp_A10IMG03").Item("ContentType")
filepathname2 = UploadRequest.Item("cp_A10IMG03").Item("FileName")
FileName2 = Right(filepathname2, Len(filepathname2) - InStrRev(filepathname2, "\"))
Value2 = UploadRequest.Item("cp_A10IMG03").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero2 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile2 = ScriptObject.CreateTextFile(Server.mappath("turismo") & "\" & FileName2)
For i = 1 To LenB(Value2)
MyFile2.Write Chr(AscB(MidB(Value2, i, 1)))
Next
MyFile2.Close

Errou
va_A10IMG03 = FileName2
Else
va_A10IMG03 = va_A10IMG03h
End If

'===============================================
If va_A11IMG04 <> "" Then
ContentType3 = UploadRequest.Item("cp_A11IMG04").Item("ContentType")
filepathname3 = UploadRequest.Item("cp_A11IMG04").Item("FileName")
FileName3 = Right(filepathname3, Len(filepathname3) - InStrRev(filepathname3, "\"))
Value3 = UploadRequest.Item("cp_A11IMG04").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero3 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile3 = ScriptObject.CreateTextFile(Server.mappath("turismo") & "\" & FileName3)
For i = 1 To LenB(Value3)
MyFile3.Write Chr(AscB(MidB(Value3, i, 1)))
Next
MyFile3.Close

Errou
va_A11IMG04 = FileName3
Else
va_A11IMG04 = va_A11IMG04h
End If

   sSel = sSel & "UPDATE turismo SET "
   sSel = sSel & "A01CATEG = '" & va_A01CATEG & "', A02NOME = '" & va_A02NOME & "', A03END = '" & va_A03END & "', A04FONE = '" & va_A04FONE & "', A05EMAIL = '" & va_A05EMAIL & "', A06WEB = '" & va_A06WEB & "', A07DESC = '" & va_A07DESC & "', A08IMG01 = '" & va_A08IMG01 & "', A09IMG02 = '" & va_A09IMG02 & "', A10IMG03 = '" & va_A10IMG03 & "', A11IMG04 = '" & va_A11IMG04 & "' "
   sSel = sSel & "WHERE A00IDTUR = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_turismo.asp"
Else
   sSel = sSel & "UPDATE turismo SET "
   sSel = sSel & "A01CATEG = '" & va_A01CATEG & "', A02NOME = '" & va_A02NOME & "', A03END = '" & va_A03END & "', A04FONE = '" & va_A04FONE & "', A05EMAIL = '" & va_A05EMAIL & "', A06WEB = '" & va_A06WEB & "', A07DESC = '" & va_A07DESC & "' "
   sSel = sSel & "WHERE A00IDTUR = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_turismo.asp"
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

