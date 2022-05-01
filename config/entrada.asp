<%
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "default.asp?mes=2"
%>
<!--#include file="_cnx.asp"-->
<!--#include file="../config.asp"-->
<html>

<head>
<title>menu</title>
<meta name="generator" content="Namo WebEditor">
<link type="text/css" rel="stylesheet" href="../suport/suport.css">
<script language="JavaScript" src="../suport/suport.js"></script>
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<table width="100%" cellpadding="0" cellspacing="0">
    <tr>
        <td width="153" align="center" valign="top"><!--#include file="menu_lateral.asp" --></td>
        <td align="center" valign="top" class="Cabecalho_Calendario" width="816">
		<br><br><br>
<% If Request("mes") = "1" Then %>
        <img src="images/atencao.gif" alt="" border="0"><br><br><span class="titulo"><font color="#FF0000">Você não tem permissão para administrar este módulo!<br><br><br><a href="entrada.asp"><img src="../imagens/botao_voltar.gif" alt="" border="0"></a></font></span>
<% Else %>
		<img src="images/gnove.gif" alt="" border="0"><br>Seja Bem-Vindo à administração do portal da<br><%= iTitulo %>
<% End If %>
		</td>
        <td align="center" valign="top" class="Cabecalho_Calendario" bgcolor="#ECECEC" width="115">
<iframe src="http://www.gnove.com.br/precam/utilidades/util.asp" name="util" id="util" width="150" height="600" scrolling="no" frameborder="0"></iframe>
		</td>
    </tr>
</table>
<p>&nbsp;</p>
</body>

</html>