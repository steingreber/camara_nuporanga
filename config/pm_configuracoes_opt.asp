<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.pagemaster.tk   \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_configuracoes_opt.asp
'CRIADO EM.........: 25/11/2006 12:28:42
'-------------------------------------------------------------------------------
'On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - CONFIGURAÇÕES"
ElseIf acao = "2" Then
   sNome = "EDITANDO AS CONFIGURAÇÕES GERAIS DO PORTAL"
End If
%>
<!--#include file="_cnx.asp"-->
<%
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"

	SQLPass = "Select * From login Where a02email = '" & Session("usuario") & "'"
	Set objRS = objConect.Execute(SQLPass)
	If Mid(objRS("a05permissoes"),18,1) = "0" Then
		response.redirect "entrada.asp?mes=1"
	End If
%>
<html>
<head>
<title>.::: PageMaster 2.0 :::.</title>
<meta name="generator" content="PageMaster v2.0"/>
<script language="JavaScript" src="lu.js"></script>
<!-- <script language=JavaScript src="picker.js"></script> -->

<SCRIPT language=Javascript>

function go2url() {
window.location = "entrada.asp";
}
function Validator(theForm)
{
  if (theForm.cp_a12tipo && !EW_hasValue(theForm.cp_a12tipo, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_a12tipo, "RADIO", "Selecione um valor para TIPO DE ENTIDADE!"))
            return false;
    }
if (theForm.cp_a05cidade && !EW_hasValue(theForm.cp_a05cidade, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a05cidade, "TEXT", "Informe um valor para o campo CIDADE"))
            return false;
    }
if (theForm.cp_a006estado && !EW_hasValue(theForm.cp_a006estado, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a006estado, "TEXT", "Informe um valor para o campo ESTADO"))
            return false;
    }
if (theForm.cp_a04endereco && !EW_hasValue(theForm.cp_a04endereco, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a04endereco, "TEXT", "Informe um valor para o campo ENDERECO"))
            return false;
    }
if (theForm.cp_a07cep && !EW_hasValue(theForm.cp_a07cep, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a07cep, "TEXT", "Informe um valor para o campo CEP"))
            return false;
    }
if (theForm.cp_a08telefone && !EW_hasValue(theForm.cp_a08telefone, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a08telefone, "TEXT", "Informe um valor para o campo TELEFONE"))
            return false;
    }
if (theForm.cp_a13fax && !EW_hasValue(theForm.cp_a13fax, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a13fax, "TEXT", "Informe um valor para o campo FAX"))
            return false;
    }
if (theForm.cp_a10email && !EW_hasValue(theForm.cp_a10email, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a10email, "TEXT", "Informe um valor para o campo EMAIL"))
            return false;
    }
if (theForm.cp_a09caixapostal && !EW_hasValue(theForm.cp_a09caixapostal, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a09caixapostal, "TEXT", "Informe um valor para o campo CAIXA POSTAL"))
            return false;
    }
if (theForm.cp_a11nivercity && !EW_hasValue(theForm.cp_a11nivercity, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a11nivercity, "TEXT", "Informe um valor para o campo ANIVERSÁRIO DA CIDADE"))
            return false;
    }
  if (theForm.cp_a15tipodeimagem && !EW_hasValue(theForm.cp_a15tipodeimagem, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_a15tipodeimagem, "RADIO", "Selecione um valor para TIPO DE IMAGEM!"))
            return false;
    }
if (theForm.cp_a02titulodapagina && !EW_hasValue(theForm.cp_a02titulodapagina, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a02titulodapagina, "TEXT", "Informe um valor para o campo TITULO DA PáGINA"))
            return false;
    }
if (theForm.cp_a01corsite && !EW_hasValue(theForm.cp_a01corsite, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a01corsite, "TEXT", "Selecione a cor do PORTAL"))
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
<!-- <script language=JavaScript src="popcalendar.js"></script> -->
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
            <table width="602" border="1" cellspacing="0" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
                <tr>
                    <td>
<%
Dim sSel, sNot
If acao = "1" Then %>
                            <table align="center" border=1 cellspacing=0 width="602" bordercolordark="white" bordercolorlight="#C0C0C0">
                        <form action="pm_configuracoes_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">TIPO DE ENTIDADE<br></span>
                    <input type="radio" name="cp_a12tipo" value="Prefeitura"><span class="fontTD">Prefeitura</span><br>
                    <input type="radio" name="cp_a12tipo" value="Câmara"><span class="fontTD">Câmara</span><br>
       </P></TD>
       <TD><P><span class="aFontePT">CIDADE<br></span><INPUT class=ud_caixa style="WIDTH:250px" name=cp_a05cidade></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">ESTADO<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a006estado></P></TD>
       <TD><P><span class="aFontePT">ENDEREÇO<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a04endereco></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CEP<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a07cep></P></TD>
       <TD><P><span class="aFontePT">TELEFONE<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a08telefone></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">FAX<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a13fax></P></TD>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a10email></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CAIXA POSTAL<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a09caixapostal></P></TD>
       <TD><P><span class="aFontePT">ANIVERSÁRIO DA CIDADE<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a11nivercity></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">TIPO DE IMAGEM<br></span>
                    <input type="radio" name="cp_a15tipodeimagem" value="1"><span class="fontTD">FLASH</span><br>
                    <input type="radio" name="cp_a15tipodeimagem" value="0"><span class="fontTD">JPEG</span><br>
       </P></TD>
       <TD><P><span class="aFontePT">IMAGEM DO TOPO DO SITE (tamanho 778x133)<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a14imagemdotitulo type="file"></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">TITULO DA PáGINA<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a02titulodapagina></P></TD>
       <TD><P><span class="aFontePT">COR DO MENU<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a01corsite></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">COR DO QUADRO<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a02corquadro></P></TD>
       <TD><P><span class="aFontePT">COR DE ENTRADA<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a03corentrada></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">COR DE SAÍDA<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a04corsaida></P></TD>
       <TD><P><span class="aFontePT">COR DO TOPO DO MENU<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a05cortopo></P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">COR DA BASE DO MENU<br></span><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a06corbase></P></td>
</TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
sNot = request.querystring("sF")
sSel = "Select * From configuracoes Where a00idconfig =" & sNot
'Set objRS = objConect.Execute(sSel)
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a12tipo                 = Trim(objRS("a12tipo"))
va_a05cidade               = Trim(objRS("a05cidade"))
va_a006estado               = Trim(objRS("a006estado"))
va_a04endereco             = Trim(objRS("a04endereco"))
va_a07cep                  = Trim(objRS("a07cep"))
va_a08telefone             = Trim(objRS("a08telefone"))
va_a13fax                  = Trim(objRS("a13fax"))
va_a10email                = Trim(objRS("a10email"))
va_a09caixapostal          = Trim(objRS("a09caixapostal"))
va_a11nivercity            = Trim(objRS("a11nivercity"))
va_a15tipodeimagem         = Trim(objRS("a15tipodeimagem"))
va_a14imagemdotitulo       = Trim(objRS("a14imagemdotitulo"))
va_a07cep                  = Trim(objRS("a07cep"))
va_a01corsite              = Trim(objRS("a01corsite"))
va_a02titulodapagina       = Trim(objRS("a02titulodapagina"))
va_a03nome                 = Trim(objRS("a03nome"))
va_a04endereco             = Trim(objRS("a04endereco"))
va_a05cidade               = Trim(objRS("a05cidade"))
va_a06estado               = Trim(objRS("a006estado"))
va_a16ligislacao           = Trim(objRS("a16ligislacao"))
va_a17tv                   = Trim(objRS("a17tv"))
va_a18noticias             = Trim(objRS("a18noticias"))
va_a19clima                = Trim(objRS("a19clima"))
va_a20enquete              = Trim(objRS("a20enquete"))
va_a21bannerG              = Trim(objRS("a21bannerG"))
va_a22bannerP              = Trim(objRS("a22bannerP"))
va_a23url                  = Trim(objRS("a23url"))
va_a24Template             = Trim(objRS("a24Template"))
va_a25altura               = Trim(objRS("a25altura"))
va_a26largura              = Trim(objRS("a26largura"))
%>
                            <table width="602" border="1" cellspacing="0" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
                        <form action="pm_configuracoes_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
       <INPUT type="hidden" name="hcp_a14imagemdotitulo" value="<%=va_a14imagemdotitulo%>">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
                                <tr>
       <TD><P><span class="aFontePT">TIPO DE ENTIDADE<br></span>
                    <input type="radio" name="cp_a12tipo" value="Prefeitura"<% If va_a12tipo = "Prefeitura" Then %> checked<% End If %>><span class="fontTD">PREFEITURA</span>&nbsp;
                    <input type="radio" name="cp_a12tipo" value="Câmara"<% If va_a12tipo = "Câmara" Then %> checked<% End If %>><span class="fontTD">CÂMARA</span><br>
       </P></TD>
       <TD><P><span class="aFontePT">TITULO DA PáGINA<br></span><INPUT class=ud_caixa style="WIDTH:295px" value='<% = va_a02titulodapagina %>' name=cp_a02titulodapagina></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CIDADE/ESTADO<br></span><INPUT class=ud_caixa style="WIDTH:200px" value='<% = va_a05cidade %>' name=cp_a05cidade>-<INPUT class=ud_caixa style="WIDTH:40px" value='<% = va_a006estado %>' name=cp_a006estado></P></TD>
       <TD><P><span class="aFontePT">ENDEREÇO<br></span><INPUT class=ud_caixa style="WIDTH:295px" value='<% = va_a04endereco %>' name=cp_a04endereco></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CEP<br></span><INPUT class=ud_caixa style="WIDTH:160px" value='<% = va_a07cep %>' name=cp_a07cep></P></TD>
       <TD><P><span class="aFontePT">TELEFONE<br></span><INPUT class=ud_caixa style="WIDTH:160px" value='<% = va_a08telefone %>' name=cp_a08telefone></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">FAX<br></span><INPUT class=ud_caixa style="WIDTH:160px" value='<% = va_a13fax %>' name=cp_a13fax></P></TD>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT class=ud_caixa style="WIDTH:260px" value='<% = va_a10email %>' name=cp_a10email></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CAIXA POSTAL<br></span><INPUT class=ud_caixa style="WIDTH:160px" value='<% = va_a09caixapostal %>' name=cp_a09caixapostal></P></TD>
       <TD><P><span class="aFontePT">ANIVERSÁRIO DA CIDADE<br></span><INPUT class=ud_caixa style="WIDTH:100px" value='<% = va_a11nivercity %>' name=cp_a11nivercity></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">TIPO DE IMAGEM DO TOPO DO PORTAL<br></span>
                    <input type="radio" name="cp_a15tipodeimagem" value="1"<% If va_a15tipodeimagem = "1" Then %> checked<% End If %>><span class="fontTD">FLASH</span>&nbsp;
                    <input type="radio" name="cp_a15tipodeimagem" value="0"<% If va_a15tipodeimagem = "0" Then %> checked<% End If %>><span class="fontTD">JPEG</span><br>
       </P></TD>
       <TD><P><span class="aFontePT">IMAGEM DO TOPO DO PORTAL - <a href="javascript:fotoZoom('<%= "imgTopo/" & va_a14imagemdotitulo%>')">VER AQUIVO</a></span><br><INPUT class=ud_caixa style="WIDTH:295px" name=cp_a14imagemdotitulo type="file"></P>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">INFORME A ALTURA E LARGURA DA IMAGEM (SOMENTE PARA ARQUIVO FLASH)<br>
	   ALTURA&nbsp;<INPUT class=ud_caixa style="WIDTH:60px" value='<% = va_a25altura %>' name=cp_a25altura>&nbsp;-&nbsp;LARGURA&nbsp;<INPUT class=ud_caixa style="WIDTH:60px" value='<% = va_a26largura %>' name=cp_a26largura></span>
	   </P></td>
</TR>
<TR>
       <TD><P><span class="aFontePT">ENDEREÇO DA TV CÂMARA (atualmente TV Senando)<br></span><INPUT class=ud_caixa style="WIDTH:295px" value='<% = va_a17tv %>' name=cp_a17tv></P></TD>
       <td colspan="2"><P><span class="aFontePT">QUANTIDADE DE NOTÍCIAS NA PRIMEIRA PÁGINA<br></span>
	   <select name="cp_a18noticias" style="WIDTH:60px" class=ud_caixa >
	   <option value="<% = va_a18noticias %>" selected><% = va_a18noticias %></option>
		<option value="5">5</option>
		<option value="7">7</option>
		<option value="9">9</option>
		<option value="11">11</option>
		<option value="13">13</option>
		</select></P>
		</td>
</TR>
<tr>
       <TD><P><span class="aFontePT">EXIBIR CLIMA TEMPO?<br></span>
           <input type="radio" name="cp_a19clima" value="1"<% If Trim(va_a19clima) = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
           <input type="radio" name="cp_a19clima" value="0"<% If Trim(va_a19clima) = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>
       </P></TD>
       <TD><P><span class="aFontePT">EXIBIR ENQUETE?<br></span>
           <input type="radio" name="cp_a20enquete" value="1"<% If Trim(va_a20enquete) = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
           <input type="radio" name="cp_a20enquete" value="0"<% If Trim(va_a20enquete) = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>
       </P></TD>
</tr>
<tr>
       <TD><P><span class="aFontePT">EXIBIR BANNERS GRANDES? 400x60<br></span>
           <input type="radio" name="cp_a21bannerG" value="1"<% If Trim(va_a21bannerG) = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
           <input type="radio" name="cp_a21bannerG" value="0"<% If Trim(va_a21bannerG) = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>
       </P></TD>
       <TD><P><span class="aFontePT">EXIBIR BANNERS PEQUENOS? 130x54<br></span>
           <input type="radio" name="cp_a22bannerP" value="1"<% If Trim(va_a22bannerP) = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
           <input type="radio" name="cp_a22bannerP" value="0"<% If Trim(va_a22bannerP) = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>
       </P></TD>
</tr>
<tr>
       <td colspan="2"><P><span class="aFontePT">ENDREÇO DO PORTAL (Exemplo: http://www.gnove.com.br)<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_a23url %>' name="cp_a23url"></P></td>
</tr>
<tr>
       <td colspan="2"><P><span class="aFontePT">LEGISLAÇÃO<br></span><INPUT class=ud_caixa style="WIDTH:295px" value='<% = va_a16ligislacao %>' name=cp_a16ligislacao></P></td>
</tr>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From configuracoes Where a00idconfig =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_configuracoes.asp"

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
va_a12tipo                 = Trim(UploadRequest.Item("cp_a12tipo").Item("Value"))
va_a05cidade               = Trim(UploadRequest.Item("cp_a05cidade").Item("Value"))
va_a006estado              = Trim(UploadRequest.Item("cp_a006estado").Item("Value"))
va_a04endereco             = Trim(UploadRequest.Item("cp_a04endereco").Item("Value"))
va_a07cep                  = Trim(UploadRequest.Item("cp_a07cep").Item("Value"))
va_a08telefone             = Trim(UploadRequest.Item("cp_a08telefone").Item("Value"))
va_a13fax                  = Trim(UploadRequest.Item("cp_a13fax").Item("Value"))
va_a10email                = Trim(UploadRequest.Item("cp_a10email").Item("Value"))
va_a09caixapostal          = Trim(UploadRequest.Item("cp_a09caixapostal").Item("Value"))
va_a11nivercity            = Trim(UploadRequest.Item("cp_a11nivercity").Item("Value"))
va_a15tipodeimagem         = Trim(UploadRequest.Item("cp_a15tipodeimagem").Item("Value"))
va_a14imagemdotitulo       = Trim(UploadRequest.Item("cp_a14imagemdotitulo").Item("Value"))
va_a02titulodapagina       = Trim(UploadRequest.Item("cp_a02titulodapagina").Item("Value"))
'va_a01corsite              = Trim(UploadRequest.Item("cp_a01corsite").Item("Value"))
va_a16ligislacao           = Trim(UploadRequest.Item("cp_a16ligislacao").Item("Value"))
va_a17tv                   = Trim(UploadRequest.Item("cp_a17tv").Item("Value"))
va_a18noticias             = Trim(UploadRequest.Item("cp_a18noticias").Item("Value"))
va_a19clima                = Trim(UploadRequest.Item("cp_a19clima").Item("Value"))
va_a20enquete              = Trim(UploadRequest.Item("cp_a20enquete").Item("Value"))
va_a21bannerG              = Trim(UploadRequest.Item("cp_a21bannerG").Item("Value"))
va_a22bannerP              = Trim(UploadRequest.Item("cp_a22bannerP").Item("Value"))
va_a23url                  = Trim(UploadRequest.Item("cp_a23url").Item("Value"))
'va_a24Template             = Trim(UploadRequest.Item("cp_a24Template").Item("Value"))
va_a25altura               = Trim(UploadRequest.Item("cp_a25altura").Item("Value"))
va_a26largura              = Trim(UploadRequest.Item("cp_a26largura").Item("Value"))

If sTipo = "1" Then
sNot = request.querystring("sF")
'===============================================
If va_a14imagemdotitulo <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM configuracoes WHERE a14imagemdotitulo = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a14imagemdotitulo").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a14imagemdotitulo").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM configuracoes WHERE a14imagemdotitulo = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a14imagemdotitulo").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("imgtopo") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a14imagemdotitulo = FileName0
End If

   sSel = sSel & "INSERT INTO configuracoes(a12tipo, a05cidade, a006estado, a04endereco, a07cep, a08telefone, a13fax, a10email, a09caixapostal, a11nivercity, a15tipodeimagem, a14imagemdotitulo, a02titulodapagina, a01corsite, a17tv, a16ligislacao, a18noticias, a19clima, a20enquete, a21bannerG, a21bannerP, a23url, a25altura, a26largura)"
   sSel = sSel & "VALUES ('" & va_a12tipo & "', '" & va_a05cidade & "', '" & va_a006estado & "', '" & va_a04endereco & "', '" & va_a07cep & "', '" & va_a08telefone & "', '" & va_a13fax & "', '" & va_a10email & "', '" & va_a09caixapostal & "', '" & va_a11nivercity & "', '" & va_a15tipodeimagem & "', '" & va_a14imagemdotitulo & "', '" & va_a02titulodapagina & "', '" & va_a01corsite & "', '" & a17tv & "', '" & a16ligislacao & "'" & a18noticias & "'" & a19clima & "'" & a20enquete & "'" & a21bannerG & "'" & a22bannerP & "'" & a23url & "'" & a25altura & "'" & a26largura & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_configuracoes.asp"

ElseIf sTipo = "2" Then

'***********************************************
va_a14imagemdotituloh       = Trim(UploadRequest.Item("hcp_a14imagemdotitulo").Item("Value"))
va_a14imagemdotituloh       = Replace(va_a14imagemdotituloh,chr(13),"<BR>" ,1)
'***********************************************
sNot = request.querystring("sF")
'===============================================
If va_a14imagemdotitulo <> "" Then
Set selectfig = objConect.Execute("SELECT * FROM configuracoes WHERE a14imagemdotitulo = '" & FileName & "';")
ContentType0 = UploadRequest.Item("cp_a14imagemdotitulo").Item("ContentType")
filepathname0 = UploadRequest.Item("cp_a14imagemdotitulo").Item("FileName")
FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
Set selectfig = objConect.Execute("SELECT * FROM configuracoes WHERE a14imagemdotitulo = '" & FileName & "';")
Value0 = UploadRequest.Item("cp_a14imagemdotitulo").Item("Value")
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
numero0 = instrrev(Request.servervariables("Path_Info"), "/")
Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("imgtopo") & "\" & FileName0)
For i = 1 To LenB(Value0)
MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
Next
MyFile0.Close

Errou
va_a14imagemdotitulo = FileName0
Else
va_a14imagemdotitulo = va_a14imagemdotituloh
End If

   sSel = sSel & "UPDATE configuracoes SET "
   sSel = sSel & "a12tipo = '" & va_a12tipo & "', a05cidade = '" & va_a05cidade & "', a006estado = '" & va_a006estado & "', a04endereco = '" & va_a04endereco & "', a07cep = '" & va_a07cep & "', a08telefone = '" & va_a08telefone & "', a13fax = '" & va_a13fax & "', a10email = '" & va_a10email & "', a09caixapostal = '" & va_a09caixapostal & "', a11nivercity = '" & va_a11nivercity & "', a15tipodeimagem = '" & va_a15tipodeimagem & "', a14imagemdotitulo = '" & va_a14imagemdotitulo & "', a02titulodapagina = '" & va_a02titulodapagina & "', a17tv = '" & va_a17tv & "', a18noticias = '" & va_a18noticias & "', a19clima = '" & va_a19clima & "', a20enquete = '" & va_a20enquete & "', a21bannerG = '" & va_a21bannerG & "', a22bannerP = '" & va_a22bannerP & "', a23url = '" & va_a23url & "', a25altura = '" & va_a25altura & "', a26largura = '" & va_a26largura & "' "
   sSel = sSel & "WHERE a00idconfig = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "entrada.asp"
Else
   sSel = sSel & "UPDATE configuracoes SET "
   sSel = sSel & "a12tipo = '" & va_a12tipo & "', a05cidade = '" & va_a05cidade & "', a006estado = '" & va_a006estado & "', a04endereco = '" & va_a04endereco & "', a07cep = '" & va_a07cep & "', a08telefone = '" & va_a08telefone & "', a13fax = '" & va_a13fax & "', a10email = '" & va_a10email & "', a09caixapostal = '" & va_a09caixapostal & "', a11nivercity = '" & va_a11nivercity & "', a15tipodeimagem = '" & va_a15tipodeimagem & "', a02titulodapagina = '" & va_a02titulodapagina & "', a17tv = '" & va_a17tv & "', a18noticias = '" & va_a18noticias & "', a19clima = '" & va_a19clima & "', a20enquete = '" & va_a20enquete & "', a21bannerG = '" & va_a21bannerG & "', a22bannerP = '" & va_a22bannerP & "', a23url = '" & va_a23url & "', a25altura = '" & va_a25altura & "', a26largura = '" & va_a26largura & "' "
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
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" align="center" cellpadding="0" cellspacing="0" width="571" bordercolordark="#FFFFFF" bordercolorlight="#C0C0C0">
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

