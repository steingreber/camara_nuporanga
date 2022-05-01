<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.5 \ ------>
'<----- /   http://www.gnove.com.br    \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_CADASTROS_opt.asp
'CRIADO EM.........: 27/2/2007 11:02:20
'-------------------------------------------------------------------------------
'On Error Resume Next
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - CADASTROS"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - CADASTROS"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>CADASTROS - PageMaster 2.0 :::.</title>
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
  if (theForm.cp_A19DATACAD && !EW_hasValue(theForm.cp_A19DATACAD, "TEXT" )) {
          if (!EW_onError(theForm, theForm.cp_A19DATACAD, "TEXT", "Informe a data correta para o campo DATA CADASTRO"))
              return false;
      }
  if (theForm.cp_A19DATACAD && !EW_checkeurodate(theForm.cp_A19DATACAD.value)) {
      if (!EW_onError(theForm, theForm.cp_A19DATACAD, "TEXT", "Data inválida para o campo DATA CADASTRO"))
          return false;
      }
if (theForm.cp_A01NOME && !EW_hasValue(theForm.cp_A01NOME, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A01NOME, "TEXT", "Informe um valor para o campo NOME"))
            return false;
    }
if (theForm.cp_A02ENDERECO && !EW_hasValue(theForm.cp_A02ENDERECO, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A02ENDERECO, "TEXT", "Informe um valor para o campo ENDERECO"))
            return false;
    }
if (theForm.cp_A03NUMERO && !EW_hasValue(theForm.cp_A03NUMERO, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A03NUMERO, "TEXT", "Informe um valor para o campo NUMERO"))
            return false;
    }
if (theForm.cp_A04BAIRRO && !EW_hasValue(theForm.cp_A04BAIRRO, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A04BAIRRO, "TEXT", "Informe um valor para o campo BAIRRO"))
            return false;
    }
if (theForm.cp_A06CIDADE && !EW_hasValue(theForm.cp_A06CIDADE, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A06CIDADE, "TEXT", "Informe um valor para o campo CIDADE"))
            return false;
    }
if (theForm.cp_A07ESTADO && !EW_hasValue(theForm.cp_A07ESTADO, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A07ESTADO, "TEXT", "Informe um valor para o campo ESTADO"))
            return false;
    }
if (theForm.cp_A23CEP && !EW_hasValue(theForm.cp_A23CEP, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A23CEP, "TEXT", "Informe um valor para o campo CEP"))
            return false;
    }
if (theForm.cp_A08FONE && !EW_hasValue(theForm.cp_A08FONE, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A08FONE, "TEXT", "Informe um valor para o campo FONE"))
            return false;
    }
if (theForm.cp_A09CELULAR && !EW_hasValue(theForm.cp_A09CELULAR, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A09CELULAR, "TEXT", "Informe um valor para o campo CELULAR"))
            return false;
    }
if (theForm.cp_A10EMAIL && !EW_hasValue(theForm.cp_A10EMAIL, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A10EMAIL, "TEXT", "Informe um valor para o campo EMAIL"))
            return false;
    }
if (theForm.cp_A17SENHA && !EW_hasValue(theForm.cp_A17SENHA, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A17SENHA, "TEXT", "Informe um valor para o campo SENHA"))
            return false;
    }
if (theForm.cp_A11CPF && !EW_hasValue(theForm.cp_A11CPF, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A11CPF, "TEXT", "Informe um valor para o campo CPF"))
            return false;
    }
if (theForm.cp_A12RG && !EW_hasValue(theForm.cp_A12RG, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A12RG, "TEXT", "Informe um valor para o campo RG"))
            return false;
    }
if (theForm.cp_A13ESTADOCIVIL && !EW_hasValue(theForm.cp_A13ESTADOCIVIL, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_A13ESTADOCIVIL, "SELECT", "Selecione um valor para o campo ESTADO CIVIL"))
            return false;
    }
  if (theForm.cp_A14SEXO && !EW_hasValue(theForm.cp_A14SEXO, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_A14SEXO, "RADIO", "Selecione um valor para SEXO!"))
            return false;
    }
if (theForm.cp_A15PROFISSAO && !EW_hasValue(theForm.cp_A15PROFISSAO, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_A15PROFISSAO, "TEXT", "Informe um valor para o campo PROFISSAO"))
            return false;
    }
  if (theForm.cp_A18DATANASC && !EW_hasValue(theForm.cp_A18DATANASC, "TEXT" )) {
          if (!EW_onError(theForm, theForm.cp_A18DATANASC, "TEXT", "Informe a data correta para o campo DATA NASCIMENTO"))
              return false;
      }
  if (theForm.cp_A18DATANASC && !EW_checkeurodate(theForm.cp_A18DATANASC.value)) {
      if (!EW_onError(theForm, theForm.cp_A18DATANASC, "TEXT", "Data inválida para o campo DATA NASCIMENTO"))
          return false;
      }
  if (theForm.cp_A20LIBERADO && !EW_hasValue(theForm.cp_A20LIBERADO, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_A20LIBERADO, "RADIO", "Selecione um valor para LIBERADO?!"))
            return false;
    }
  if (theForm.cp_A21LIBERADOATEH && !EW_hasValue(theForm.cp_A21LIBERADOATEH, "TEXT" )) {
          if (!EW_onError(theForm, theForm.cp_A21LIBERADOATEH, "TEXT", "Informe a data correta para o campo LIBERADO ATE..."))
              return false;
      }
  if (theForm.cp_A21LIBERADOATEH && !EW_checkeurodate(theForm.cp_A21LIBERADOATEH.value)) {
      if (!EW_onError(theForm, theForm.cp_A21LIBERADOATEH, "TEXT", "Data inválida para o campo LIBERADO ATE..."))
          return false;
      }
if (theForm.cp_A22PLANO && !EW_hasValue(theForm.cp_A22PLANO, "SELECT" )) {
        if (!EW_onError(theForm, theForm.cp_A22PLANO, "SELECT", "Selecione um valor para o campo PLANO"))
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
                            <table align="center" border=1 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_CADASTROS_opt.asp?acao=4&tp=1" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">DATA CADASTRO<br></span><INPUT value="<%=date%>" class=ud_caixa style="WIDTH:150px" name=cp_A19DATACAD>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" align="absmiddle" onClick="popUpCalendar(this, this.form.cp_A19DATACAD,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A01NOME></P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">ENDEREÇO/NÚMERO<br></span><INPUT value="" class=ud_caixa style="WIDTH:89%" name=cp_A02ENDERECO>/<INPUT value="" class=ud_caixa style="WIDTH:10%" name=cp_A03NUMERO></P></td>
</TR>
<TR>
       <TD><P><span class="aFontePT">BAIRRO<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A04BAIRRO></P></TD>
       <TD><P><span class="aFontePT">COMPLEMENTO<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A05COMPLEMENTO></P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">CIDADE/ESTADO<br></span><INPUT value="" class=ud_caixa style="WIDTH:94%" name=cp_A06CIDADE>/<INPUT value="" class=ud_caixa style="WIDTH:5%" name=cp_A07ESTADO></P></td>
</TR>
<TR>
       <TD><P><span class="aFontePT">CEP<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A23CEP></P></TD>
       <TD><P><span class="aFontePT">FONE<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A08FONE></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CELULAR<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A09CELULAR></P></TD>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A10EMAIL></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">SENHA<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A17SENHA></P></TD>
       <TD><P><span class="aFontePT">CPF<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A11CPF></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">RG<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A12RG></P></TD>
       <TD><P><span class="aFontePT">ESTADO CIVIL<br></span>
           <select name="cp_A13ESTADOCIVIL" size="1" class="ud_caixa">
                <option selected>Selecione</option>
                    <option value="CASADO">CASADO</option>
                    <option value="SOLTEIRO">SOLTEIRO</option>
                    <option value="VIUVO">VIUVO</option>
                    <option value="DESQUITADO">DESQUITADO</option>
                    <option value="DIVORCIADO">DIVORCIADO</option>
                    <option value="UNIAO ESTAVEL">UNIAO ESTAVEL</option>
                </select>
</TR>
<TR>
       <TD><P><span class="aFontePT">SEXO&nbsp;</span>
                    <input type="radio" name="cp_A14SEXO" value="M"><span class="fontTD">MASCULINO</span>&nbsp;
                    <input type="radio" name="cp_A14SEXO" value="F"><span class="fontTD">FEMININO</span>&nbsp;
       </P></TD>
       <TD><P><span class="aFontePT">PROFISSAO<br></span><INPUT value="" class=ud_caixa style="WIDTH:100%" name=cp_A15PROFISSAO></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">DATA NASCIMENTO<br></span><INPUT value="" class=ud_caixa style="WIDTH:150px" name=cp_A18DATANASC>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" align="absmiddle" onClick="popUpCalendar(this, this.form.cp_A18DATANASC,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">LIBERADO?&nbsp;</span>
                    <input type="radio" name="cp_A20LIBERADO" value="1"><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_A20LIBERADO" value="0" checked><span class="fontTD">NÃO</span>&nbsp;
       </P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">LIBERADO ATE...<br></span><INPUT value="" class=ud_caixa style="WIDTH:150px" name=cp_A21LIBERADOATEH>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" align="absmiddle" onClick="popUpCalendar(this, this.form.cp_A21LIBERADOATEH,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P></P>
</TR>
<TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
sNot = request.querystring("sF")
sSel = "Select * From CADASTROS Where A00ID =" & sNot
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_A19DATACAD              = Trim(objRS("A19DATACAD"))
va_A01NOME                 = Trim(objRS("A01NOME"))
va_A02ENDERECO             = Trim(objRS("A02ENDERECO"))
va_A03NUMERO               = Trim(objRS("A03NUMERO"))
va_A04BAIRRO               = Trim(objRS("A04BAIRRO"))
va_A05COMPLEMENTO          = Trim(objRS("A05COMPLEMENTO"))
va_A06CIDADE               = Trim(objRS("A06CIDADE"))
va_A07ESTADO               = Trim(objRS("A07ESTADO"))
va_A23CEP                  = Trim(objRS("A23CEP"))
va_A08FONE                 = Trim(objRS("A08FONE"))
va_A09CELULAR              = Trim(objRS("A09CELULAR"))
va_A10EMAIL                = Trim(objRS("A10EMAIL"))
va_A17SENHA                = Trim(objRS("A17SENHA"))
va_A11CPF                  = Trim(objRS("A11CPF"))
va_A12RG                   = Trim(objRS("A12RG"))
va_A13ESTADOCIVIL          = Trim(objRS("A13ESTADOCIVIL"))
va_A14SEXO                 = Trim(objRS("A14SEXO"))
va_A15PROFISSAO            = Trim(objRS("A15PROFISSAO"))
va_A18DATANASC             = Trim(objRS("A18DATANASC"))
va_A20LIBERADO             = Trim(objRS("A20LIBERADO"))
va_A21LIBERADOATEH         = Trim(objRS("A21LIBERADOATEH"))
'-----------------------------------------------
%>
                            <table align="center" border=1 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="#808080">
                        <form action="pm_CADASTROS_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" enctype="multipart/form-data" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">DATA CADASTRO<br></span><INPUT class=ud_caixa style="WIDTH:150px" value='<% = va_A19DATACAD %>' name=cp_A19DATACAD>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" align="absmiddle" onClick="popUpCalendar(this, this.form.cp_A19DATACAD,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A01NOME %>' name=cp_A01NOME></P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">ENDEREÇO/NUMERO<br></span><INPUT class=ud_caixa style="WIDTH:79%" value='<% = va_A02ENDERECO %>' name=cp_A02ENDERECO>/<INPUT class=ud_caixa style="WIDTH:20%" value='<% = va_A03NUMERO %>' name=cp_A03NUMERO></P></td>
</TR>
<TR>
       <TD><P><span class="aFontePT">BAIRRO<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A04BAIRRO %>' name=cp_A04BAIRRO></P></TD>
       <TD><P><span class="aFontePT">COMPLEMENTO<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A05COMPLEMENTO %>' name=cp_A05COMPLEMENTO></P></TD>
</TR>
<TR>
       <td colspan="2"><P><span class="aFontePT">CIDADE/ESTADO<br></span><INPUT class=ud_caixa style="WIDTH:89%" value='<% = va_A06CIDADE %>' name=cp_A06CIDADE>/<INPUT class=ud_caixa style="WIDTH:10%" value='<% = va_A07ESTADO %>' name=cp_A07ESTADO></P></td>
</TR>
<TR>
       <TD><P><span class="aFontePT">CEP<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A23CEP %>' name=cp_A23CEP></P></TD>
       <TD><P><span class="aFontePT">FONE<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A08FONE %>' name=cp_A08FONE></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">CELULAR<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A09CELULAR %>' name=cp_A09CELULAR></P></TD>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A10EMAIL %>' name=cp_A10EMAIL></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">SENHA<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A17SENHA %>' name=cp_A17SENHA></P></TD>
       <TD><P><span class="aFontePT">CPF<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A11CPF %>' name=cp_A11CPF></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">RG<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A12RG %>' name=cp_A12RG></P></TD>
       <TD><P><span class="aFontePT">ESTADO CIVIL<br></span>
           <select name="cp_A13ESTADOCIVIL" size="1" class="ud_caixa">
		   			<option>Selecione</option>
                <option value="<%=va_A13ESTADOCIVIL%>" selected><%=va_A13ESTADOCIVIL%></option>
                    <option value="CASADO">CASADO</option>
                    <option value="SOLTEIRO">SOLTEIRO</option>
                    <option value="VIUVO">VIUVO</option>
                    <option value="DESQUITADO">DESQUITADO</option>
                    <option value="DIVORCIADO">DIVORCIADO</option>
                    <option value="UNIAO ESTAVEL">UNIAO ESTAVEL</option>
                </select>
</TR>
<TR>
       <TD><P><span class="aFontePT">SEXO</span>
                    <input type="radio" name="cp_A14SEXO" value="M"<% If va_A14SEXO = "M" Then %> checked<% End If %>><span class="fontTD">MASCULINO</span>&nbsp;
                    <input type="radio" name="cp_A14SEXO" value="F"<% If va_A14SEXO = "F" Then %> checked<% End If %>><span class="fontTD">FEMININO</span>
       </P></TD>
       <TD><P><span class="aFontePT">PROFISSÃO<br></span><INPUT class=ud_caixa style="WIDTH:100%" value='<% = va_A15PROFISSAO %>' name=cp_A15PROFISSAO></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">DATA NASCIMENTO<br></span><INPUT class=ud_caixa style="WIDTH:150px" value='<% = va_A18DATANASC %>' name=cp_A18DATANASC>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" align="absmiddle" onClick="popUpCalendar(this, this.form.cp_A18DATANASC,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P><span class="aFontePT">LIBERADO?&nbsp;</span>
                    <input type="radio" name="cp_A20LIBERADO" value="1"<% If va_A20LIBERADO = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_A20LIBERADO" value="0"<% If va_A20LIBERADO = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>
       </P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">LIBERADO ATÉ...<br></span><INPUT class=ud_caixa style="WIDTH:150px" value='<% = va_A21LIBERADOATEH %>' name=cp_A21LIBERADOATEH>&nbsp;<input type="image" src="images/pm_calendar.gif" alt="Pick a Date" align="absmiddle" onClick="popUpCalendar(this, this.form.cp_A21LIBERADOATEH,'dd/mm/yyyy');return false;"></P></TD>
       <TD><P></P></TD>
</TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From CADASTROS Where A00ID =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_CADASTROS.asp"
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
va_A19DATACAD              = Trim(UploadRequest.Item("cp_A19DATACAD").Item("Value"))
va_A01NOME                 = Trim(UploadRequest.Item("cp_A01NOME").Item("Value"))
va_A02ENDERECO             = Trim(UploadRequest.Item("cp_A02ENDERECO").Item("Value"))
va_A03NUMERO               = Trim(UploadRequest.Item("cp_A03NUMERO").Item("Value"))
va_A04BAIRRO               = Trim(UploadRequest.Item("cp_A04BAIRRO").Item("Value"))
va_A05COMPLEMENTO          = Trim(UploadRequest.Item("cp_A05COMPLEMENTO").Item("Value"))
va_A06CIDADE               = Trim(UploadRequest.Item("cp_A06CIDADE").Item("Value"))
va_A07ESTADO               = Trim(UploadRequest.Item("cp_A07ESTADO").Item("Value"))
va_A23CEP                  = Trim(UploadRequest.Item("cp_A23CEP").Item("Value"))
va_A08FONE                 = Trim(UploadRequest.Item("cp_A08FONE").Item("Value"))
va_A09CELULAR              = Trim(UploadRequest.Item("cp_A09CELULAR").Item("Value"))
va_A10EMAIL                = Trim(UploadRequest.Item("cp_A10EMAIL").Item("Value"))
va_A17SENHA                = Trim(UploadRequest.Item("cp_A17SENHA").Item("Value"))
va_A11CPF                  = Trim(UploadRequest.Item("cp_A11CPF").Item("Value"))
va_A12RG                   = Trim(UploadRequest.Item("cp_A12RG").Item("Value"))
va_A13ESTADOCIVIL          = Trim(UploadRequest.Item("cp_A13ESTADOCIVIL").Item("Value"))
va_A14SEXO                 = Trim(UploadRequest.Item("cp_A14SEXO").Item("Value"))
va_A15PROFISSAO            = Trim(UploadRequest.Item("cp_A15PROFISSAO").Item("Value"))
va_A18DATANASC             = Trim(UploadRequest.Item("cp_A18DATANASC").Item("Value"))
va_A20LIBERADO             = Trim(UploadRequest.Item("cp_A20LIBERADO").Item("Value"))
va_A21LIBERADOATEH         = Trim(UploadRequest.Item("cp_A21LIBERADOATEH").Item("Value"))
'-----------------------------------------------

If sTipo = "1" Then
sNot = request.querystring("sF")
   sSel = sSel & "INSERT INTO CADASTROS(A19DATACAD, A01NOME, A02ENDERECO, A03NUMERO, A04BAIRRO, A05COMPLEMENTO, A06CIDADE, A07ESTADO, A23CEP, A08FONE, A09CELULAR, A10EMAIL, A17SENHA, A11CPF, A12RG, A13ESTADOCIVIL, A14SEXO, A15PROFISSAO, A18DATANASC, A20LIBERADO, A21LIBERADOATEH)"
   sSel = sSel & "VALUES ('" & va_A19DATACAD & "', '" & va_A01NOME & "', '" & va_A02ENDERECO & "', '" & va_A03NUMERO & "', '" & va_A04BAIRRO & "', '" & va_A05COMPLEMENTO & "', '" & va_A06CIDADE & "', '" & va_A07ESTADO & "', '" & va_A23CEP & "', '" & va_A08FONE & "', '" & va_A09CELULAR & "', '" & va_A10EMAIL & "', '" & va_A17SENHA & "', '" & va_A11CPF & "', '" & va_A12RG & "', '" & va_A13ESTADOCIVIL & "', '" & va_A14SEXO & "', '" & va_A15PROFISSAO & "', '" & va_A18DATANASC & "', '" & va_A20LIBERADO & "', '" & va_A21LIBERADOATEH & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CADASTROS.asp"

ElseIf sTipo = "2" Then

sNot = request.querystring("sF")
   sSel = sSel & "UPDATE CADASTROS SET "
   sSel = sSel & "A19DATACAD = '" & va_A19DATACAD & "', A01NOME = '" & va_A01NOME & "', A02ENDERECO = '" & va_A02ENDERECO & "', A03NUMERO = '" & va_A03NUMERO & "', A04BAIRRO = '" & va_A04BAIRRO & "', A05COMPLEMENTO = '" & va_A05COMPLEMENTO & "', A06CIDADE = '" & va_A06CIDADE & "', A07ESTADO = '" & va_A07ESTADO & "', A23CEP = '" & va_A23CEP & "', A08FONE = '" & va_A08FONE & "', A09CELULAR = '" & va_A09CELULAR & "', A10EMAIL = '" & va_A10EMAIL & "', A17SENHA = '" & va_A17SENHA & "', A11CPF = '" & va_A11CPF & "', A12RG = '" & va_A12RG & "', A13ESTADOCIVIL = '" & va_A13ESTADOCIVIL & "', A14SEXO = '" & va_A14SEXO & "', A15PROFISSAO = '" & va_A15PROFISSAO & "', A18DATANASC = '" & va_A18DATANASC & "', A20LIBERADO = '" & va_A20LIBERADO & "', A21LIBERADOATEH = '" & va_A21LIBERADOATEH & "' "
   sSel = sSel & "WHERE A00ID = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CADASTROS.asp"
Else
   sSel = sSel & "UPDATE CADASTROS SET "
   sSel = sSel & "A19DATACAD = '" & va_A19DATACAD & "', A01NOME = '" & va_A01NOME & "', A02ENDERECO = '" & va_A02ENDERECO & "', A03NUMERO = '" & va_A03NUMERO & "', A04BAIRRO = '" & va_A04BAIRRO & "', A05COMPLEMENTO = '" & va_A05COMPLEMENTO & "', A06CIDADE = '" & va_A06CIDADE & "', A07ESTADO = '" & va_A07ESTADO & "', A23CEP = '" & va_A23CEP & "', A08FONE = '" & va_A08FONE & "', A09CELULAR = '" & va_A09CELULAR & "', A10EMAIL = '" & va_A10EMAIL & "', A17SENHA = '" & va_A17SENHA & "', A11CPF = '" & va_A11CPF & "', A12RG = '" & va_A12RG & "', A13ESTADOCIVIL = '" & va_A13ESTADOCIVIL & "', A14SEXO = '" & va_A14SEXO & "', A15PROFISSAO = '" & va_A15PROFISSAO & "', A18DATANASC = '" & va_A18DATANASC & "', A20LIBERADO = '" & va_A20LIBERADO & "', A21LIBERADOATEH = '" & va_A21LIBERADOATEH & "' "
   sSel = sSel & "WHERE A00ID = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_CADASTROS.asp"
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

