<%
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "default.asp?mes=2"
%>
<!--#include file="_cnx.asp"-->
<!--#include file="../config.asp"-->
<html>

<head>
<title>No title</title>
<meta name="generator" content="Namo WebEditor">
<link type="text/css" rel="stylesheet" href="pagemaster.css">
<base target="main"></head>

<body bgcolor="white" text="black" link="blue" vlink="blue" alink="blue" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<table width="100%" height="13" cellpadding="0" cellspacing="0">
    <tr>
        <td height="13" bgcolor="#ECECEC" align="center"><a href="http://www.gnove.com.br" target="_blank"><img src="images/logo_gnove.gif" width="83" height="30" border="0"></a></td>
        <td height="13" align="center" bgcolor="#ECECEC" class="fontTD">Olá, <%= Session("nome") %></td>
        <td height="13" align="center" bgcolor="#ECECEC" class="fontTD">Portal: <%= iTitulo %></td>
        <td height="13" align="center" bgcolor="#ECECEC" class="fontTD"><a href="vaza.asp" target="_parent"><img src="images/topo_sair.png" width="19" height="19" border="0" align="absmiddle"></a><a href="vaza.asp" target="_parent">&nbsp;SAIR</a></td>
    </tr>
    <tr>
        <td width="180" height="13" align="center" background="images/fundo_esq.gif">
            <p class="fontTD"><font color="white">»»&nbsp;</font><a href="../default.asp" target="_blank"><font color="white">acessar portal </font></a><font color="white">««</font></p>
        </td>
        <td height="13" align="center" bgcolor="#ECECEC" class="fontTD" colspan="3" background="images/fundo_top.gif">&nbsp;</td>
    </tr>
</table>
</body>

</html>