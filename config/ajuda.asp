<!--#include file="_cnx.asp"-->
<html>

<head>
<title>Ajuda PreCam</title>
<meta name="robots" content="none">
<link type="text/css" rel="stylesheet" href="pagemaster.css">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<table width="248" height="322" cellpadding="0" cellspacing="0">
    <tr>
        <td width="124" height="25" bgcolor="#ECECEC">
            <p style="margin-left:5px;"><img src="images/logo_gnove.gif" width="83" height="30" border="0"></p>
        </td>
        <td width="124" height="25" bgcolor="#ECECEC" align="center"><a href="javascript:close();"><img src="../imagens/botao_fechar.gif" width="88" height="16" border="0"></a></td>
    </tr>
    <tr>
        <td width="238" height="293" valign="top" bgcolor="#D1FDE0" colspan="2">
<%
sNot = request("sF")
sSel = "Select * From ajuda Where a00id =" & sNot
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sSel,objConect,3,3
If objRS.BOF = False AND objRS.EOF = False Then
%>	<br>
            <p class="aFontePT" style="margin-right:5px; margin-left:5px;"><b><%= objRS(1) %></b><br><br>
<span class="subtitulo"><%= objRS(2) %></span>
			</p>
<%
Else
%>
            <p class="aFontePT" style="margin-right:5px; margin-left:5px;"><br><br><br><b>Não há item relacionado a este tópico!</b></p>
<%
End If
objRS.Close
%>
        </td>
    </tr>
</table>
</body>

</html>