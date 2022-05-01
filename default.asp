<%Response.ContentType = "text/html; charset=iso-8859-1"%>

<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<!--#include file="data.asp"-->
<html>

<head>
<title><%= iTitulo & " - " & iLegis %></title>
<meta name="generator" content="Namo WebEditor">
<link type="text/css" rel="stylesheet" href="suport/suport_def.css">
<script language="JavaScript" src="suport/suport.js"></script>
</head>

<body bgcolor="white" text="black" link="blue" vlink="<%= cCorTopoBase %>" alink="red" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<table width="100%" cellpadding="0" cellspacing="0" height="423">
                <TR>
<% If tipoimg = "0" Then %>
                    <td width="100%" colspan="3" align="center"><img src="config/imgtopo/<%= imgtopo %>" alt="" border="0"></td>
<% Else %>
                    <TD colspan="3" width="100%" align="center"><script>ConteudoFlash('config/imgTopo/<%= imgtopo %>','<%= iAltura %>','<%= iLargura %>');</script></TD>
<% End If %>
                </TR>
    <tr>
        <td colspan="3" height="5"></td>
    </tr>
    <tr>
        <td colspan="3" bgcolor="<%= cCorTopoBase %>">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="79"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a1.gif" width="69" height="40" border="0"></td>
                    <td width="46"><b><a href="default.asp"><font size="2" face="Arial" color="white">Home</font></a></b></td>
                    <td width="10" valign="baseline"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a2.gif" border="0"></td>
                    <td width="98"><b><a href="#" onClick="ajax('contato.asp','conteudo')"><font size="2" face="Arial" color="white">Fale Conosco</font></a></b></td>
                    <td width="10" valign="bottom"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a2.gif" border="0"></td>
                    <td width="174"><!-- <b><a href="#"onClick="ajax('calendario.asp','conteudo')"><font size="2" face="Arial" color="white">Calendário de Obrigações</font></a></b> --></td>
                    <td width="10"></td>
                    <td align="right"><font size="2" face="Arial" color="white"><b><%= hoje %></b></font></td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td width="153" height="272" rowspan="2" valign="top">
<!--#include file="menu_lateral.asp"-->
        </td>
        <td height="20">
<%
If Session("iUser") = True Then
	If Session("xSexo") = "M" Then
%>
		&nbsp;&nbsp;<font size="2" face="Arial" color="#999999"><b><%= mensagemtempo %></b>, </font><font size="2" face="Arial" color="#999999"><%= Session("vUser") %>&nbsp;-&nbsp;<a href="logar_vaza.asp">(X)Sair</a></font>
<%
	Else
%>
		&nbsp;&nbsp;<font size="2" face="Arial" color="#999999"><b><%= mensagemtempo %></b>, </font><font size="2" face="Arial" color="#999999"><%= Session("vUser") %>&nbsp;-&nbsp;<a href="logar_vaza.asp">(X)Sair</a></font>
<%
	End If
Else
%>
		&nbsp;&nbsp;<font size="2" face="Arial" color="#999999"><b><%= mensagemtempo %></b>, </font><a href="#" onClick="ajax('logar.asp','conteudo')"><font size="2" face="Arial" color="#999999">cadastre-se aqui ou faça seu login</font></a>
<%
End If
%>

		</td>
        <td height="272" rowspan="2" align="right" valign="top">
<!--#include file="lateral_direita.asp"-->
        </td>
    </tr>
    <tr>
        <td height="225" align="center" valign="top">
<% If iBannerG = "1" Then %>
<!--#include file="banner.asp" -->
<% End If %>
<p align="left" style="margin-right:6px; margin-left:6px;">
<div align="left" id="conteudo" style="width: 98%;">
<!--#include file="fotos_entrada.asp"-->
<!--#include file="noticias.asp"-->
</div>
</p>
        </td>
    </tr>
    <tr>
        <td valign="top">&nbsp;</td>
        <td>&nbsp;</td>
        <td align="right"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a10.gif" width="119" height="47" border="0"></td>
    </tr>
    <tr>
        <td width="937" bgcolor="<%= cCorTopoBase %>" valign="bottom" colspan="2" height="41">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="79">&nbsp;</td>
                    <td width="73"><b><a href="default.asp"><font size="2" face="Arial" color="white">Home</font></a></b></td>
                    <td width="10" valign="bottom"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a2.gif" border="0"></td>
                    <td width="125"><b><a href="#" onClick="ajax('contato.asp','conteudo')"><font size="2" face="Arial" color="white">Fale Conosco</font></a></b></td>
                    <td width="8" valign="bottom"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a2.gif" border="0"></td>
                    <td width="643"><!-- <p><b><a href="#"onClick="ajax('calendario.asp','conteudo')"><font size="2" face="Arial" color="white">Calendário de Obrigações</font></a></b></p> --></td>
                </tr>
            </table>
        </td>
        <td align="right" bgcolor="<%= cCorTopoBase %>" height="41"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a9.gif" width="119" height="54" border="0"></td>
    </tr>
    <tr>
        <td height="41" colspan="3">
            <table width="100%" cellpadding="0" cellspacing="0" height="67">
                <tr>
                    <td align="center" height="67"><a href="http://www.gnove.com.br/visitas.asp?id=8" target="_blank"><img src="imagens/poweredby.jpg" width="152" height="44" border="0"></a></td>
                    <td align="center" height="67">
                        <p><b><font size="1" face="Arial"><%= iTitulo & " " & iLegis %></font></b></p>
                    </td>
                    <td align="center" height="67"><font size="2" face="Arial"><b><!--#include file="contador.asp"--></b></font></td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<p>&nbsp;</p>
</body>

</html>