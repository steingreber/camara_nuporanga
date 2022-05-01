<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<html>

<head>
<title><%= iTitulo & " " & iLegis %></title>
<meta name="author" content="Alex L. Steingreber">
<meta name="robots" content="none">
<link type="text/css" rel="stylesheet" href="suport/suport_def.css">
<script language="JavaScript" src="suport/suport.js"></script>
</head>

<body bgcolor="<%= ca02corTopoTabela %>" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000" leftmargin="5" topmargin="5" marginwidth="5">
<table width="299" cellpadding="0" cellspacing="0">
    <tr>
        <td width="299" height="238" align="center" valign="middle" bgcolor="<%= ca03corCorpoTabela %>">
            <table cellpadding="0" cellspacing="0" width="246" height="114">
            <form name="indique">
                <tr>
                    <td width="246" height="38" align="center" class="tittext"><strong><%= iTitulo & " " & iLegis %></strong></td>
                </tr>
                <tr>
                    <td width="246" class="tittext" height="38"><strong>Lembrar senha.</strong><br></td>
                </tr>
                    <tr>
                        <td width="246" class="navtext">Digite seu email de cadastro</td>
                    </tr>
                    <tr>
                        <td width="246"><input type="text" name="formtext1" class="ud_caixa" size="37"></td>
                    </tr>
                <tr>
                    <td width="246" align="center" height="31" valign="bottom"><input type="image" name="formimage1" src="imagens/botao_senha.gif" border="0"></td>
                </tr>
            </form>				
            </table>
        </td>
    </tr>
</table>
</body>

</html>
