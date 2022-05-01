<%
'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.gnove.com.br    \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: pm_login_opt.asp
'CRIADO EM.........: 18/12/2006 12:45:52
'-------------------------------------------------------------------------------
'On Error Resume Next
Dim acao, sNome
acao = Request.QueryString("acao")
If acao = "1" Then
   sNome = "NOVO REGISTRO - PERMISSÕES DE ACESSO"
ElseIf acao = "2" Then
   sNome = "EDITAR REGISTRO - PERMISSÕES DE ACESSO"
End If
%>
<!--#include file="_cnx.asp"-->
<html>
<head>
<title>LOGIN - PageMaster 2.0 :::.</title>
<meta name="generator" content="PageMaster v2.0"/>
<script language="JavaScript" src="validator.js"></script>
<script language=JavaScript src="popcalendar.js"></script>
<script language="Javascript1.2" src="editor.js"></script>
<script>_editor_url = "";</script>
<SCRIPT language=Javascript>

function go2url() {
window.location = "pm_login.asp";
}
function Validator(theForm)
{
if (theForm.cp_a01nome && !EW_hasValue(theForm.cp_a01nome, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a01nome, "TEXT", "Informe um valor para o campo NOME"))
            return false;
    }
if (theForm.cp_a02email && !EW_hasValue(theForm.cp_a02email, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a02email, "TEXT", "Informe um valor para o campo EMAIL"))
            return false;
    }
if (theForm.cp_a03senha && !EW_hasValue(theForm.cp_a03senha, "TEXT" )) {
        if (!EW_onError(theForm, theForm.cp_a03senha, "TEXT", "Informe um valor para o campo SENHA"))
            return false;
    }
  if (theForm.cp_a04ativo && !EW_hasValue(theForm.cp_a04ativo, "RADIO" )) {
          if (!EW_onError(theForm, theForm.cp_a04ativo, "RADIO", "Selecione um valor para ATIVO!"))
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
        <td height="5" class="tabelaitem">
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
                            <table align="center" border=0 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_login_opt.asp?acao=4&tp=1" method="post" name="frmInc" id="frmInc" onSubmit="return Validator(this);">
                           <input type="hidden" name="tp" value="1">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT value="" class=ud_caixa style="WIDTH:290px" name=cp_a01nome></P></TD>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT value="" class=ud_caixa style="WIDTH:290px" name=cp_a02email></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">SENHA<br></span><input type="password" name="cp_a03senha" class="ud_caixa" style="WIDTH:100px"></P></TD>
       <TD><P><span class="aFontePT">ESTE USUÁRIO ESTA ATIVO?<br></span>
                    <input type="radio" name="cp_a04ativo" value="1"><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_a04ativo" value="0"><span class="fontTD">NÃO</span>
       </P></TD>
</TR>
                            <TR>
                                <td colspan="2">
                                    <span class="aFontePT">PERMISSÕES (SELECIONE O QUE ESTE USUÁRIO PODER ADMINISTRAR)</span><br>
                                    <table width="100%" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes23"><strong><font color="#FF0000">ACESSO AO ADMINISTRADOR</font></strong></td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes02">NOTÍCIAS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes03">VEREADORES</td>
                                         </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes04">LEGISLAÇÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes05">COMISSÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes06">SEÇÕES PLENÁRIAS</td>
                                        </tr>
                                        <tr>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes07">FOTOS DO MUNICÍPIO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes08">MESA DIRETORA</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes09">CONCURSOS PÚBLICOS</td>
                                         </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes10">LICITAÇÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes11">CONTAS PÚBLICAS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes12">LINKS</td>
                                        </tr>
                                        <tr>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes13">FOTOS PRIMEIRA PÁGINA</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes14">TURISMO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes15">BANNERS</td>
                                         </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes16">SECRETARIAS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes17">UTILIDADES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes18">CONFIGURAÇOES</td>
                                        </tr>
                                        <tr>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes19">CONTRATOS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes20">AUDIENCIAS PÚBLICAS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes21">INTRANET</td>
                                        </tr>
										<tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes22">PERMISSÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes01">SIMBOLOS DA CIDADE</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes24">TELEFONES ÚTEIS</td>
                                        </tr>
										<tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes25">CLIMA TEMPO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes26">CONVITES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes27">TOMADA DE PREÇOS</td>
                                        </tr>
										<tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes28">CONCORRÊNCIA PÚBLICA</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes29">PREGÃO PRESENCIAL</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes30">PREGÃO ELETRÔNICO</td>
                                        </tr>
										<tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes31">LEILÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes32">DISPESAS DE LICITAÇÃO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes33">INEXIGIBILIDADE DE LICITAÇÃO</td>
                                        </tr>
										<tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes34">DADOS DO MUNICÍPIO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes35">ADMINISTRAÇÃO DO MUNICÍPIO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes36">FALA PREFEITO</td>
                                        </tr>
										<tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes37">HISTÓRIA DA CIDADE</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes38">MODELOS DO PORTAL</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes39">SUB_CATEGORIAS</td>
                                        </tr>
                                    </table>
                                </td>
                            </TR>
                        </form>    
                            </table>
<% ElseIf acao = "2" Then
sNot = request.querystring("sF")
sSel = "Select * From login Where a00id =" & sNot
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
'-----------------------------------------------
va_a01nome         = Trim(objRS("a01nome"))
va_a02email        = Trim(objRS("a02email"))
va_a03senha        = Trim(objRS("a03senha"))
va_a04ativo        = Trim(objRS("a04ativo"))
va_a05permissoes01 = Mid(Trim(objRS("a05permissoes")),1,1)
va_a05permissoes02 = Mid(Trim(objRS("a05permissoes")),2,1)
va_a05permissoes03 = Mid(Trim(objRS("a05permissoes")),3,1)
va_a05permissoes04 = Mid(Trim(objRS("a05permissoes")),4,1)
va_a05permissoes05 = Mid(Trim(objRS("a05permissoes")),5,1)
va_a05permissoes06 = Mid(Trim(objRS("a05permissoes")),6,1)
va_a05permissoes07 = Mid(Trim(objRS("a05permissoes")),7,1)
va_a05permissoes08 = Mid(Trim(objRS("a05permissoes")),8,1)
va_a05permissoes09 = Mid(Trim(objRS("a05permissoes")),9,1)
va_a05permissoes10 = Mid(Trim(objRS("a05permissoes")),10,1)
va_a05permissoes11 = Mid(Trim(objRS("a05permissoes")),11,1)
va_a05permissoes12 = Mid(Trim(objRS("a05permissoes")),12,1)
va_a05permissoes13 = Mid(Trim(objRS("a05permissoes")),13,1)
va_a05permissoes14 = Mid(Trim(objRS("a05permissoes")),14,1)
va_a05permissoes15 = Mid(Trim(objRS("a05permissoes")),15,1)
va_a05permissoes16 = Mid(Trim(objRS("a05permissoes")),16,1)
va_a05permissoes17 = Mid(Trim(objRS("a05permissoes")),17,1)
va_a05permissoes18 = Mid(Trim(objRS("a05permissoes")),18,1)
va_a05permissoes19 = Mid(Trim(objRS("a05permissoes")),19,1)
va_a05permissoes20 = Mid(Trim(objRS("a05permissoes")),20,1)
va_a05permissoes21 = Mid(Trim(objRS("a05permissoes")),21,1)
va_a05permissoes22 = Mid(Trim(objRS("a05permissoes")),22,1)
va_a05permissoes23 = Mid(Trim(objRS("a05permissoes")),23,1)
va_a05permissoes24 = Mid(Trim(objRS("a05permissoes")),24,1)
va_a05permissoes25 = Mid(Trim(objRS("a05permissoes")),25,1)
va_a05permissoes26 = Mid(Trim(objRS("a05permissoes")),26,1)
va_a05permissoes27 = Mid(Trim(objRS("a05permissoes")),27,1)
va_a05permissoes28 = Mid(Trim(objRS("a05permissoes")),28,1)
va_a05permissoes29 = Mid(Trim(objRS("a05permissoes")),29,1)
va_a05permissoes30 = Mid(Trim(objRS("a05permissoes")),30,1)
va_a05permissoes31 = Mid(Trim(objRS("a05permissoes")),31,1)
va_a05permissoes32 = Mid(Trim(objRS("a05permissoes")),32,1)
va_a05permissoes33 = Mid(Trim(objRS("a05permissoes")),33,1)
va_a05permissoes34 = Mid(Trim(objRS("a05permissoes")),34,1)
va_a05permissoes35 = Mid(Trim(objRS("a05permissoes")),35,1)
va_a05permissoes36 = Mid(Trim(objRS("a05permissoes")),36,1)
va_a05permissoes37 = Mid(Trim(objRS("a05permissoes")),37,1)
va_a05permissoes38 = Mid(Trim(objRS("a05permissoes")),38,1)
va_a05permissoes39 = Mid(Trim(objRS("a05permissoes")),39,1)
teste01 = Trim(objRS("a05permissoes"))

'-----------------------------------------------
%>
                            <table align="center" border=0 cellspacing=0 width="100%" bordercolordark="white" bordercolorlight="black">
                        <form action="pm_login_opt.asp?acao=4&tp=2&sF=<% = sNot %>" method="post" name="frmInc" id="frmInc" onSubmit="return Validator(this);">
                               <tr>
                                    <td colspan="2">
                                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">&nbsp;<input type="submit" name="frmId" class="botao" value=".::Gravar::." height=19></td>
                                </tr>
<TR>
       <TD><P><span class="aFontePT">NOME<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a01nome %>' name=cp_a01nome></P></TD>
       <TD><P><span class="aFontePT">EMAIL<br></span><INPUT class=ud_caixa style="WIDTH:290px" value='<% = va_a02email %>' name=cp_a02email></P></TD>
</TR>
<TR>
       <TD><P><span class="aFontePT">SENHA<br></span><input type="password" name="cp_a03senha" value="<% = va_a03senha %>" class="ud_caixa" style="WIDTH:290px"></P></TD>
       <TD><P><span class="aFontePT">ESTE USUÁRIO ESTA ATIVO?<br></span>
                    <input type="radio" name="cp_a04ativo" value="1" <% If va_a04ativo = "1" Then %> checked<% End If %>><span class="fontTD">SIM</span>&nbsp;
                    <input type="radio" name="cp_a04ativo" value="0" <% If va_a04ativo = "0" Then %> checked<% End If %>><span class="fontTD">NÃO</span>
       </P></TD>
</TR>
                            <TR>
                                <td colspan="2">
                                    <span class="aFontePT">PERMISSÕES (SELECIONE O QUE ESTE USUÁRIO PODER ADMINISTRAR)</span><br>
                                    <table width="100%" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes23"<% If va_a05permissoes23 = "1" Then %> checked<% End If %>><strong><font color="#FF0000">ACESSO AO ADMINISTRADOR</font></strong></td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes02"<% If va_a05permissoes02 = "1" Then %> checked<% End If %>>NOTÍCIAS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes03"<% If va_a05permissoes03 = "1" Then %> checked<% End If %>>VEREADORES</td>
										</TR>
										<TR>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes04"<% If va_a05permissoes04 = "1" Then %> checked<% End If %>>LEGISLAÇÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes05"<% If va_a05permissoes05 = "1" Then %> checked<% End If %>>COMISSÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes06"<% If va_a05permissoes06 = "1" Then %> checked<% End If %>>SEÇÕES PLENÁRIAS</td>
                                        </tr>
                                        <tr>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes07"<% If va_a05permissoes07 = "1" Then %> checked<% End If %>>FOTOS DO MUNICÍPIO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes08"<% If va_a05permissoes08 = "1" Then %> checked<% End If %>>MESA DIRETORA</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes09"<% If va_a05permissoes09 = "1" Then %> checked<% End If %>>CONCURSOS PÚBLICOS</td>
                                        </tr>
										<tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes10"<% If va_a05permissoes10 = "1" Then %> checked<% End If %>>LICITAÇÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes11"<% If va_a05permissoes11 = "1" Then %> checked<% End If %>>CONTAS PÚBLICAS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes12"<% If va_a05permissoes12 = "1" Then %> checked<% End If %>>LINKS</td>
                                        </tr>
                                        <tr>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes13"<% If va_a05permissoes13 = "1" Then %> checked<% End If %>>FOTOS PRIMEIRA PÁGINA</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes14"<% If va_a05permissoes14 = "1" Then %> checked<% End If %>>TURISMO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes15"<% If va_a05permissoes15 = "1" Then %> checked<% End If %>>BANNERS</td>
                                         </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes16"<% If va_a05permissoes16 = "1" Then %> checked<% End If %>>SECRETARIAS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes17"<% If va_a05permissoes17 = "1" Then %> checked<% End If %>>UTILIDADES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes18"<% If va_a05permissoes18 = "1" Then %> checked<% End If %>>CONFIGURAÇOES</td>
                                        </tr>
                                        <tr>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes19"<% If va_a05permissoes19 = "1" Then %> checked<% End If %>>CONTRATOS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes20"<% If va_a05permissoes20 = "1" Then %> checked<% End If %>>AUDIENCIAS PÚBLICAS</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes21"<% If va_a05permissoes21 = "1" Then %> checked<% End If %>>ÁREA RESTRITA</td>
                                         </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes22"<% If va_a05permissoes22 = "1" Then %> checked<% End If %>>PERMISSÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes01"<% If va_a05permissoes01 = "1" Then %> checked<% End If %>>SIMBOLOS DA CIDADE</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes24"<% If va_a05permissoes24 = "1" Then %> checked<% End If %>>TELEFONES ÚTEIS</td>
                                        </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes25"<% If va_a05permissoes25 = "1" Then %> checked<% End If %>>CLIMA TEMPO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes26"<% If va_a05permissoes26 = "1" Then %> checked<% End If %>>CONVITES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes27"<% If va_a05permissoes27 = "1" Then %> checked<% End If %>>TOMADA DE PREÇOS</td>
                                        </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes28"<% If va_a05permissoes28 = "1" Then %> checked<% End If %>>CONCORRÊNCIA PÚBLICA</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes29"<% If va_a05permissoes29 = "1" Then %> checked<% End If %>>PREGÃO PRESENCIAL</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes30"<% If va_a05permissoes30 = "1" Then %> checked<% End If %>>PREGÃO ELETRÔNICO</td>
                                        </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes31"<% If va_a05permissoes31 = "1" Then %> checked<% End If %>>LEILÕES</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes32"<% If va_a05permissoes32 = "1" Then %> checked<% End If %>>DISPESAS DE LICITAÇÃO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes33"<% If va_a05permissoes33 = "1" Then %> checked<% End If %>>INEXIGIBILIDADE DE LICITAÇÃO</td>
                                        </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes34"<% If va_a05permissoes34 = "1" Then %> checked<% End If %>>DADOS DO MUNICÍPIO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes35"<% If va_a05permissoes35 = "1" Then %> checked<% End If %>>ADMINISTRAÇÃO DO MUNICÍPIO</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes36"<% If va_a05permissoes36 = "1" Then %> checked<% End If %>>FALA PREFEITO</td>
                                        </tr>
										 <tr>
										    <td class="fmb"><input type="checkbox" name="cp_a05permissoes37"<% If va_a05permissoes37 = "1" Then %> checked<% End If %>>HISTÓRIA DA CIDADE</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes38"<% If va_a05permissoes38 = "1" Then %> checked<% End If %>>MODELOS DO PORTAL</td>
                                            <td class="fmb"><input type="checkbox" name="cp_a05permissoes39"<% If va_a05permissoes39 = "1" Then %> checked<% End If %>>SUB-CATEGORIAS</td>
                                        </tr>
                                    </table>
                                </TD>
                            </TR>
                        </form>    
                            </table>
<% ElseIf acao = "3" Then
sNot = request.querystring("sF")
sSel = "Delete From login Where a00id =" & sNot
Set objRS = objConect.Execute(sSel)
response.redirect "pm_login.asp"

ElseIf acao = "4" Then
sNot = request.querystring("sF")
Response.Expires = 0

Dim sTipo, sCmd, sPer(40)
sTipo  = request.querystring("tp")
If Trim(Request("cp_a05permissoes01")) = "on" Then sPer(0) = 1: Else sPer(0) = 0
If Trim(Request("cp_a05permissoes02")) = "on" Then sPer(1) = 1: Else sPer(1) = 0
If Trim(Request("cp_a05permissoes03")) = "on" Then sPer(2) = 1: Else sPer(2) = 0
If Trim(Request("cp_a05permissoes04")) = "on" Then sPer(3) = 1: Else sPer(3) = 0
If Trim(Request("cp_a05permissoes05")) = "on" Then sPer(4) = 1: Else sPer(4) = 0
If Trim(Request("cp_a05permissoes06")) = "on" Then sPer(5) = 1: Else sPer(5) = 0
If Trim(Request("cp_a05permissoes07")) = "on" Then sPer(6) = 1: Else sPer(6) = 0
If Trim(Request("cp_a05permissoes08")) = "on" Then sPer(7) = 1: Else sPer(7) = 0
If Trim(Request("cp_a05permissoes09")) = "on" Then sPer(8) = 1: Else sPer(8) = 0
If Trim(Request("cp_a05permissoes10")) = "on" Then sPer(9) = 1: Else sPer(9) = 0
If Trim(Request("cp_a05permissoes11")) = "on" Then sPer(10) = 1: Else sPer(10) = 0
If Trim(Request("cp_a05permissoes12")) = "on" Then sPer(11) = 1: Else sPer(11) = 0
If Trim(Request("cp_a05permissoes13")) = "on" Then sPer(12) = 1: Else sPer(12) = 0
If Trim(Request("cp_a05permissoes14")) = "on" Then sPer(13) = 1: Else sPer(13) = 0
If Trim(Request("cp_a05permissoes15")) = "on" Then sPer(14) = 1: Else sPer(14) = 0
If Trim(Request("cp_a05permissoes16")) = "on" Then sPer(15) = 1: Else sPer(15) = 0
If Trim(Request("cp_a05permissoes17")) = "on" Then sPer(16) = 1: Else sPer(16) = 0
If Trim(Request("cp_a05permissoes18")) = "on" Then sPer(17) = 1: Else sPer(17) = 0
If Trim(Request("cp_a05permissoes19")) = "on" Then sPer(18) = 1: Else sPer(18) = 0
If Trim(Request("cp_a05permissoes20")) = "on" Then sPer(19) = 1: Else sPer(19) = 0
If Trim(Request("cp_a05permissoes21")) = "on" Then sPer(20) = 1: Else sPer(20) = 0
If Trim(Request("cp_a05permissoes22")) = "on" Then sPer(21) = 1: Else sPer(21) = 0
If Trim(Request("cp_a05permissoes23")) = "on" Then sPer(22) = 1: Else sPer(22) = 0
If Trim(Request("cp_a05permissoes24")) = "on" Then sPer(23) = 1: Else sPer(23) = 0
If Trim(Request("cp_a05permissoes25")) = "on" Then sPer(24) = 1: Else sPer(24) = 0
If Trim(Request("cp_a05permissoes26")) = "on" Then sPer(25) = 1: Else sPer(25) = 0
If Trim(Request("cp_a05permissoes27")) = "on" Then sPer(26) = 1: Else sPer(26) = 0
If Trim(Request("cp_a05permissoes28")) = "on" Then sPer(27) = 1: Else sPer(27) = 0
If Trim(Request("cp_a05permissoes29")) = "on" Then sPer(28) = 1: Else sPer(28) = 0
If Trim(Request("cp_a05permissoes30")) = "on" Then sPer(29) = 1: Else sPer(29) = 0
If Trim(Request("cp_a05permissoes31")) = "on" Then sPer(30) = 1: Else sPer(30) = 0
If Trim(Request("cp_a05permissoes32")) = "on" Then sPer(31) = 1: Else sPer(31) = 0
If Trim(Request("cp_a05permissoes33")) = "on" Then sPer(32) = 1: Else sPer(32) = 0
If Trim(Request("cp_a05permissoes34")) = "on" Then sPer(33) = 1: Else sPer(33) = 0
If Trim(Request("cp_a05permissoes35")) = "on" Then sPer(34) = 1: Else sPer(34) = 0
If Trim(Request("cp_a05permissoes36")) = "on" Then sPer(35) = 1: Else sPer(35) = 0
If Trim(Request("cp_a05permissoes37")) = "on" Then sPer(36) = 1: Else sPer(36) = 0
If Trim(Request("cp_a05permissoes38")) = "on" Then sPer(37) = 1: Else sPer(37) = 0
If Trim(Request("cp_a05permissoes39")) = "on" Then sPer(38) = 1: Else sPer(38) = 0

'-----------------------------------------------
va_a01nome       = Trim(Request("cp_a01nome"))
va_a02email      = Trim(Request("cp_a02email"))
va_a03senha      = Trim(Request("cp_a03senha"))
va_a04ativo      = Trim(Request("cp_a04ativo"))
va_a05permissoes = sPer(0) & sPer(1) & sPer(2) & sPer(3) & sPer(4) & sPer(5)
va_a05permissoes = va_a05permissoes & sPer(6) & sPer(7) & sPer(8) & sPer(9) & sPer(10) &sPer(11)
va_a05permissoes = va_a05permissoes & sPer(12) & sPer(13) & sPer(14) & sPer(15) & sPer(16) & sPer(17)
va_a05permissoes = va_a05permissoes & sPer(18) & sPer(19) & sPer(20) & sPer(21) & sPer(22) & sPer(23)
va_a05permissoes = va_a05permissoes & sPer(24) & sPer(25) & sPer(26) & sPer(27) & sPer(28) & sPer(29)
va_a05permissoes = va_a05permissoes & sPer(30) & sPer(31) & sPer(32) & sPer(33) & sPer(34) & sPer(35)
va_a05permissoes = va_a05permissoes & sPer(36) & sPer(37) & sPer(38)
'-----------------------------------------------

If sTipo = "1" Then
sNot = request.querystring("sF")
   sSel = sSel & "INSERT INTO login(a01nome, a02email, a03senha, a04ativo, a05permissoes)"
   sSel = sSel & "VALUES ('" & va_a01nome & "', '" & va_a02email & "', '" & va_a03senha & "', '" & va_a04ativo & "', '" & va_a05permissoes & "')"
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_login.asp"

ElseIf sTipo = "2" Then

sNot = request.querystring("sF")
   sSel = sSel & "UPDATE login SET "
   sSel = sSel & "a01nome = '" & va_a01nome & "', a02email = '" & va_a02email & "', a03senha = '" & va_a03senha & "', a04ativo = '" & va_a04ativo & "', a05permissoes = '" & va_a05permissoes & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_login.asp"
Else

   sSel = sSel & "UPDATE login SET "
   sSel = sSel & "a01nome = '" & va_a01nome & "', a02email = '" & va_a02email & "', a03senha = '" & va_a03senha & "', a04ativo = '" & va_a04ativo & "', a05permissoes = '" & va_a05permissoes & "' "
   sSel = sSel & "WHERE a00id = " & sNot
   Set objRS = objConect.Execute(sSel)
   response.redirect "pm_login.asp"
End If
End If
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
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" align="center" cellpadding="0" cellspacing="0" width="571" bordercolordark="black" bordercolorlight="black">
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

