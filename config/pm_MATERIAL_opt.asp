<!--#include file="_cnx.asp"-->
<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.5 \ ------>
'<----- /   http://www.gnove.com.br    \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_MATERIAL_opt.asp
'CRIADO EM.........: 14/2/2007 09:24:14
'-------------------------------------------------------------------------------
On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
Dim acao, sNome, sCateg, NomeCateg
sCateg = Request("categ")
sPer   = Request("per")
acao   = Request.QueryString("acao")

iCateg = "Select * From CATEGORIAS Where A00IDC1="&sCateg
Set objRSct = Server.CreateObject("ADODB.Recordset")
objRSct.Open iCateg, objConect, 3, 3
NomeCateg = objRSct(1)
objRSct.Close

If acao = "1" Then
   sNome = "NOVO REGISTRO - "&NomeCateg
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - "&NomeCateg
End If
%>
<html>
<head>
<title>MATERIAL - PageMaster 2.0 :::.</title>
<meta name="generator" content="PageMaster v2.0"/>
<script language="JavaScript" src="validator.js"></script>
<script language=JavaScript src="popcalendar.js"></script>
<script language="Javascript1.2" src="editor.js"></script>
<script>_editor_url = "";</script>
<SCRIPT language=Javascript>

function go2url() {
window.location = "pm_MATERIAL.asp?categ=<%= sCateg %>&per=<%= sPer %>";
}
function Validator(theForm)
{
  if (theForm.cp_A04DATA && !EW_hasValue(theForm.cp_A04DATA, "TEXT" )) {
          if (!EW_onError(theForm, theForm.cp_A04DATA, "TEXT", "Informe a data correta para o campo DATA"))
              return false;
      }
  if (theForm.cp_A04DATA && !EW_checkeurodate(theForm.cp_A04DATA.value)) {
      if (!EW_onError(theForm, theForm.cp_A04DATA, "TEXT", "Data inválida para o campo DATA"))
          return false;
      }
if (theForm.cp_A11SUBCATEG && !EW_hasValue(theForm.cp_A11SUBCATEG, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_A11SUBCATEG, "SELECT", "Selecione um valor para o campo SUB-CATEGORIA"))
            return false;
    }	  
if (theForm.cp_A01TITULO && !EW_hasValue(theForm.cp_A01TITULO, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A01TITULO, "TEXT", "Informe um valor para o campo TITULO"))
            return false;
    }
if (theForm.cp_A05CATEG && !EW_hasValue(theForm.cp_A05CATEG, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_A05CATEG, "SELECT", "Selecione um valor para o campo CATEGORIA"))
            return false;
    }
if (theForm.cp_A06BUSCA && !EW_hasValue(theForm.cp_A06BUSCA, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A06BUSCA, "TEXT", "Informe um valor para o campo CAMPO DE BUSCA NO SITE"))
            return false;
    }
  if (theForm.cp_A07EXIBIR && !EW_hasValue(theForm.cp_A07EXIBIR, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_A07EXIBIR, "RADIO", "Selecione um valor para EXIBIR ESTE ITEM?!"))
            return false;
    }
  if (theForm.cp_A08BLOQUEADO && !EW_hasValue(theForm.cp_A08BLOQUEADO, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_A08BLOQUEADO, "RADIO", "Selecione um valor para ITEM DEPENDENTE DE CADASTRO?!"))
            return false;
    }
if (theForm.cp_A02TEXTO && !EW_hasValue(theForm.cp_A02TEXTO, "TEXTAREA" )) {
            if (!EW_onError(theForm, theForm.cp_A02TEXTO, "TEXTAREA", "Informe um valor para o campo TEXTO DA MATÉRIA"))
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
<body background="images/bg.gif" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000" leftmargin="0" topmargin="0">
<table width='100%' cellpadding='0' cellspacing='0' height='100%'>
    <tr>
        <td align='center' valign='middle'>
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
            <table width="610" border="1" cellspacing="0" align="center" bordercolorlight="#808080" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table align="center" border=1 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="#dedede">
                        <form action="pm_MATERIAL_opt.asp?acao=4&tp=1&categ=<%= sCateg %>&per=<%= sPer %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <td colspan="2"><P><span class="aFontePT">CATEGORIA<br></span>
           <select name="cp_A05CATEG" size="1" class="ud_caixa">
<%sLin = "Select * From CATEGORIAS Where A02EXIBIR=1 AND A00IDC1="&sCateg&" Order By A00IDC1"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option>Selecione</option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("A00IDC1") %>" selected><%= objRs("A00IDC1") &"-"& objRs("A01CATEG1") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
	   </td>
</tr>
<TR>
       <TD><P><span class="aFontePT">HÁ SUB-CATEGORIA PARA ESTE CADASTRO?</span><br>
                    <input type="radio" name="cp_A12HASUBCATEG" value="1" checked onClick="exibi_subcateg();"><span class="fontTD">SIM</span>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=1', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
                    <!-- <input type="radio" name="cp_A12HASUBCATEG" value="0" onclick="esconde_subcateg();"><span class="fontTD">NÃO</span>&nbsp; -->
       </P></TD>
		<td>
<!-- <div id="subcateg" style="display:none"> -->
		<span class="aFontePT">SUB CATEGORIA<br></span>
            <select name="cp_A11SUBCATEG" size="1" class="ud_caixa">
<%sLin = "Select * From SUBCATEG Where A03EXIBIR=1 AND A01CATEG="& sCateg &" Order By A01CATEG"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
            <option selected>Selecione</option>
<%Do While not objRs.EOF%>
            <option value="<%= objRs("A00IDSUB") %>"><%= objRs("A01CATEG")& "-" &objRs("A02TITULO") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
            </select>&nbsp;<a href="pm_SUBCATEG_opt.asp?acao=1&categ=<%= sCateg %>&per=<%= sPer %>"><img src="images/add-comment-red.gif" alt="ADCIONAR SUB-CATEGORIA" border="0" align="absmiddle"></a>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=1', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
<!-- </div> -->
		</td>
</TR>
<TR>
       <td><P><span class="aFontePT">DATA<br></span><INPUT value="<%=date%>" class=ud_caixa style="WIDTH:150px" name=cp_A04DATA>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" align="absmiddle" onClick="popUpCalendar(this, this.form.cp_A04DATA,'dd/mm/yyyy');return false;"></P></td>
       <td><P><span class="aFontePT">ANO<br></span><INPUT value="<%=Year(date)%>" class=ud_caixa style="WIDTH:50px" name=cp_A13ANO></P></td>

</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">TITULO (máximo 250 caracteres)&nbsp;<input type="text" name="remLen" value="250" size="3" readonly class="ud_caixa"><br></span><textarea cols="" rows="" name="cp_A01TITULO" class="ud_caixa" style="width:100%; height:40;" onKeyDown="textCounter(this.form.cp_A01TITULO,this.form.remLen,250);" onKeyUp="textCounter(this.form.cp_A01TITULO,this.form.remLen,250);"></textarea></P></td>
</TR>
<tr>
       <td colspan="2"><P><span class="aFontePT">ARQUIVO<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A10ARQUIVO type="file"></P></td>
</TR>
<TR>
       <TD><P><span class="aFontePT">ITEM DEPENDENTE DE CADASTRO?&nbsp;</span>
                    <input type="radio" name="cp_A08BLOQUEADO" value="1"><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_A08BLOQUEADO" value="0" checked><span class="fontTD">NÃO</span>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=2', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
       </P></TD>
       <TD><P><span class="aFontePT">EXIBIR ESTE ITEM?&nbsp;</span>
                    <input type="radio" name="cp_A07EXIBIR" value="1" checked><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_A07EXIBIR" value="0"><span class="fontTD">NÃO</span>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=3', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
       </P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">CAMPO DE BUSCA NO SITE<br></span><INPUT value="" class=ud_caixa style="WIDTH:95%" name=cp_A06BUSCA>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=4', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a></P></td>
</tr>
<tr>
       <td colspan="2">
           <span class="aFontePT">TEXTO DA MATÉRIA<br></span><TEXTAREA class=ud_caixa style="width:100%; height:100; border-style:none;" name=cp_A02TEXTO></textarea>
           <script language="javascript1.2">
           editor_generate('cp_A02TEXTO'); // field, width, height
           </script>
         </td>
</TR>
<TR>
       <td colspan="2">
           <span class="aFontePT">TEXTO COMPLEMENTAR<br></span><TEXTAREA class=ud_caixa style="width:100%; height:100; border-style:none;" name=cp_A03TEXTO></textarea>
           <script language="javascript1.2">
           editor_generate('cp_A03TEXTO'); // field, width, height
           </script>
         </td>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
sNot = request.querystring("sF")
sSel = "Select * From MATERIAL Where A00ID =" & sNot
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_A04DATA                 = Trim(objRS("A04DATA"))
va_A01TITULO               = Trim(objRS("A01TITULO"))
va_A05CATEG                = Trim(objRS("A05CATEG"))
va_A10ARQUIVO              = Trim(objRS("A10ARQUIVO"))
va_A06BUSCA                = Trim(objRS("A06BUSCA"))
va_A07EXIBIR               = Trim(objRS("A07EXIBIR"))
va_A08BLOQUEADO            = Trim(objRS("A08BLOQUEADO"))
va_A02TEXTO                = Trim(objRS("A02TEXTO"))
va_A03TEXTO                = Trim(objRS("A03TEXTO"))
va_A11SUBCATEG             = Trim(objRS("A11SUBCATEG"))
va_A12HASUBCATEG           = Trim(objRS("A12HASUBCATEG"))
va_A13ANO                  = Trim(objRS("A13ANO"))
'-----------------------------------------------
%>
                            <table align="center" border=1 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_MATERIAL_opt.asp?acao=4&tp=2&sF=<% = sNot %>&categ=<%= sCateg %>&per=<%= sPer %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
       <INPUT type="hidden" name="hcp_A10ARQUIVO" value="<%=va_A10ARQUIVO%>">
<tr>
       <td colspan="2">
        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19>
       </td>
</tr>
<TR>
       <TD><P><span class="aFontePT">DATA<br></span><INPUT class=ud_caixa style="WIDTH:60px" value='<% = va_A04DATA %>' name=cp_A04DATA>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Selecione a data" align="absmiddle" onClick="popUpCalendar(this, this.form.cp_A04DATA,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">ANO<br></span><INPUT class=ud_caixa style="WIDTH:50px" value='<% = va_A13ANO %>' name=cp_A13ANO></P></TD>
</TR>
<tr>
       <td colspan="2"><P><span class="aFontePT">Há SUB-CATEGORIA PARA ESTE CADASTRO?</span><br>
                    <input type="radio" name="cp_A12HASUBCATEG" value="1"<% If va_A12HASUBCATEG = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;<a href="pm_SUBCATEG_opt.asp?acao=1&categ=<%= sCateg %>"><img src="images/add-comment-red.gif" alt="ADCIONAR SUB-CATEGORIA" border="0" align="absmiddle"></a>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=5', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
                    <!-- <input type="radio" name="cp_A12HASUBCATEG" value="0"<% If va_A12HASUBCATEG = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>&nbsp; -->
       </P></TD>
</tr>
<TR>
       <td colspan="2"><P><span class="aFontePT">TITULO (máximo 250 caracteres)&nbsp;<input type="text" name="remLen" value="250" size="3" readonly class="ud_caixa"><br></span><TEXTAREA class=ud_caixa style="width:100%; height:40;" name=cp_A01TITULO onKeyDown="textCounter(this.form.cp_A01TITULO,this.form.remLen,250);" onKeyUp="textCounter(this.form.cp_A01TITULO,this.form.remLen,250);"><% = va_A01TITULO %></textarea></td>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">CATEGORIA<br></span>
           <select name="cp_A05CATEG" size="1" class="ud_caixa">
                <option selected>Selecione</option>
<% sLin = "Select * From CATEGORIAS Where A02EXIBIR=1 Order By A00IDC1"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
Do While not objRs.EOF %>
                <option value="<%= objRs("A00IDC1") %>"<% If Trim(va_A05CATEG) = Trim(objRs("A00IDC1")) Then%> selected<% End If %>><%= objRs("A00IDC1") & "-" & objRs("A01CATEG1") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
	   </td>
</tr>
<TR>
       <td colspan="2"><P><span class="aFontePT">SUB-CATEGORIA<br></span>
           <select name="cp_A11SUBCATEG" size="1" class="ud_caixa">
                <option value="" selected>Selecione</option>
<% sLin = "Select * From SUBCATEG Where A03EXIBIR=1 Order By A01CATEG"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
Do While not objRs.EOF %>
                <option value="<%= objRs("A00IDSUB") %>"<% If Trim(va_A11SUBCATEG) = Trim(objRs("A00IDSUB")) Then%> selected<% End If %>><%= objRs("A01CATEG") & "-" & objRs("A02TITULO") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>&nbsp;<a href="pm_SUBCATEG_opt.asp&categ=<%= sCateg %>&per=<%= sPer %>"><img src="images/add-comment-red.gif" alt="ADCIONAR SUB-CATEGORIA" border="0" align="absmiddle"></a>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=1', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
	   </td>
</tr>
<tr>
       <td colspan="2"><P><span class="aFontePT">ARQUIVO - <a href="javascript:fotoZoom('<%= "arquivos/" & va_A10ARQUIVO%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:100%" name=cp_A10ARQUIVO type="file"><BR></P>
</TR>
<TR>
       <TD><P><span class="aFontePT">ITEM DEPENDENTE DE CADASTRO?&nbsp;</span>
                    <input type="radio" name="cp_A08BLOQUEADO" value="1"<% If va_A08BLOQUEADO = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_A08BLOQUEADO" value="0"<% If va_A08BLOQUEADO = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=2', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
       </P></TD>
       <TD><P><span class="aFontePT">EXIBIR ESTE ITEM?&nbsp;</span>
                    <input type="radio" name="cp_A07EXIBIR" value="1"<% If va_A07EXIBIR = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_A07EXIBIR" value="0"<% If va_A07EXIBIR = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=3', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a>
       </P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">CAMPO DE BUSCA NO SITE<br></span><INPUT class=ud_caixa style="WIDTH:95%" value='<% = va_A06BUSCA %>' name=cp_A06BUSCA>&nbsp;<a href="javascript:open_window('win', 'ajuda.asp?sF=4', 50, 50, 248, 322, 0, 0, 0, 0, 0);"><img src="images/ajuda.gif" alt="Ajuda!" border="0" align="absmiddle"></a></P></td>
</tr>
<tr>
       <td colspan="2">
         <span class="aFontePT">TEXTO DA MATÉRIA<br></span><TEXTAREA class=ud_caixa style="width:100%; height:100; border-style:none;" name=cp_A02TEXTO><% = va_A02TEXTO %></textarea>
         <script language="javascript1.2">
         editor_generate('cp_A02TEXTO'); // field, width, height
         </script>
       </td>
</TR>
<TR>
       <td colspan="2">
         <span class="aFontePT">TEXTO COMPLEMENTAR<br></span><TEXTAREA class=ud_caixa style="width:100%; height:100; border-style:none;" name=cp_A03TEXTO><% = va_A03TEXTO %></textarea>
         <script language="javascript1.2">
         editor_generate('cp_A03TEXTO'); // field, width, height
         </script>
       </td>
</TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From MATERIAL Where A00ID =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_MATERIAL.asp?categ="&sCateg&"&per="&sPer
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
va_A04DATA                 = Trim(UploadRequest.Item("cp_A04DATA").Item("Value"))
va_A01TITULO               = UCase(Trim(UploadRequest.Item("cp_A01TITULO").Item("Value")))
va_A05CATEG                = Trim(UploadRequest.Item("cp_A05CATEG").Item("Value"))
va_A10ARQUIVO              = Trim(UploadRequest.Item("cp_A10ARQUIVO").Item("Value"))
va_A06BUSCA                = Trim(UploadRequest.Item("cp_A06BUSCA").Item("Value"))
va_A07EXIBIR               = Trim(UploadRequest.Item("cp_A07EXIBIR").Item("Value"))
va_A08BLOQUEADO            = Trim(UploadRequest.Item("cp_A08BLOQUEADO").Item("Value"))
va_A02TEXTO                = Trim(UploadRequest.Item("cp_A02TEXTO").Item("Value"))
va_A03TEXTO                = Trim(UploadRequest.Item("cp_A03TEXTO").Item("Value"))
va_A12HASUBCATEG           = Trim(UploadRequest.Item("cp_A12HASUBCATEG").Item("Value"))
If va_A12HASUBCATEG <> "0" Then
va_A11SUBCATEG             = Trim(UploadRequest.Item("cp_A11SUBCATEG").Item("Value"))
Else
va_A11SUBCATEG             = "0"
End If
va_A13ANO                  = Trim(UploadRequest.Item("cp_A13ANO").Item("Value"))
'-----------------------------------------------

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_A10ARQUIVO <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM MATERIAL WHERE A10ARQUIVO = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_A10ARQUIVO").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_A10ARQUIVO").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM MATERIAL WHERE A10ARQUIVO = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_A10ARQUIVO").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("arquivos") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

va_A10ARQUIVO = FileName0
End If

   sSel = sSel & "INSERT INTO MATERIAL(A04DATA, A01TITULO, A05CATEG, A10ARQUIVO, A06BUSCA, A07EXIBIR, A08BLOQUEADO, A02TEXTO, A03TEXTO, A12HASUBCATEG, A11SUBCATEG, A13ANO) "
   sSel = sSel & "VALUES ('" & va_A04DATA & "', '" & va_A01TITULO & "', '" & va_A05CATEG & "', '" & va_A10ARQUIVO & "', '" & va_A06BUSCA & "', '" & va_A07EXIBIR & "', '" & va_A08BLOQUEADO & "', '" & va_A02TEXTO & "', '" & va_A03TEXTO & "', '" & va_A12HASUBCATEG & "', '" & va_A11SUBCATEG& "', '" & va_A13ANO& "')"
   Set objRS = objConect.Execute(sSel)
response.redirect "pm_MATERIAL.asp?categ="&sCateg&"&per="&sPer

ElseIf sTipo = "2" Then

va_A10ARQUIVOh              = Trim(UploadRequest.Item("hcp_A10ARQUIVO").Item("Value"))
va_A10ARQUIVOh              = Replace(va_A10ARQUIVOh,chr(13),"<BR>" ,1)
sNot = request.querystring("sF")
'===============================================
If va_A10ARQUIVO <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM MATERIAL WHERE A10ARQUIVO = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_A10ARQUIVO").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_A10ARQUIVO").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM MATERIAL WHERE A10ARQUIVO = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_A10ARQUIVO").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("arquivos") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_A10ARQUIVO = FileName0
Else
va_A10ARQUIVO = va_A10ARQUIVOh
End If

   sSel = sSel & "UPDATE MATERIAL SET "
   sSel = sSel & "A04DATA = '" & va_A04DATA & "', A01TITULO = '" & va_A01TITULO & "', A05CATEG = '" & va_A05CATEG & "', A10ARQUIVO = '" & va_A10ARQUIVO & "', A06BUSCA = '" & va_A06BUSCA & "', A07EXIBIR = '" & va_A07EXIBIR & "', A08BLOQUEADO = '" & va_A08BLOQUEADO & "', A02TEXTO = '" & va_A02TEXTO & "', A03TEXTO = '" & va_A03TEXTO & "', A11SUBCATEG = '" & va_A11SUBCATEG & "', A12HASUBCATEG = '" & va_A12HASUBCATEG & "', A13ANO = '" & va_A13ANO & "' "
   sSel = sSel & "WHERE A00ID = " & sNot
   Set objRS = objConect.Execute(sSel)
response.redirect "pm_MATERIAL.asp?categ="&sCateg&"&per="&sPer
Else
   sSel = sSel & "UPDATE MATERIAL SET "
   sSel = sSel & "A04DATA = '" & va_A04DATA & "', A01TITULO = '" & va_A01TITULO & "', A05CATEG = '" & va_A05CATEG & "', A06BUSCA = '" & va_A06BUSCA & "', A07EXIBIR = '" & va_A07EXIBIR & "', A08BLOQUEADO = '" & va_A08BLOQUEADO & "', A02TEXTO = '" & va_A02TEXTO & "', A03TEXTO = '" & va_A03TEXTO & "', A11SUBCATEG = '" & va_A11SUBCATEG & "', A12HASUBCATEG = '" & va_A12HASUBCATEG & "', A13ANO = '" & va_A13ANO & "' "
   sSel = sSel & "WHERE A00ID = " & sNot
   Set objRS = objConect.Execute(sSel)
response.redirect "pm_MATERIAL.asp?categ="&sCateg&"&per="&sPer
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
</table></body>
</html>
<%
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
%>

