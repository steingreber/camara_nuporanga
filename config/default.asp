<!--#include file="_cnx.asp"-->
<!--#include file="../config.asp"-->
<html>
<head>
<script Language="JavaScript">
// Begin
function Validator(theForm)
{
  if (theForm.username.value == "")
  {
    alert("Digite o nome do Usuário!!");
    theForm.username.focus();
    return (false);
  }
  if (theForm.password.value == "")
  {
    alert("Digite sua Senha!!");
    theForm.password.focus();
    return (false);
  }
  return (true);
}
function placeFocus() {
if (document.forms.length > 0) {
var field = document.forms[0];
for (i = 0; i < field.length; i++) {
if ((field.elements[i].type == "text") || (field.elements[i].type == "textarea") || (field.elements[i].type.toString().charAt(0) == "s")) {
document.forms[0].elements[i].focus();
break;
         }
      }
   }
}
// End
</script>
<title>Gerenciador de Páginas GNove Web Studio</title>
<link type="text/css" rel="stylesheet" href="pagemaster.css">
</head>
<body background="images/bg.gif" onLoad="placeFocus();">

<table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0'>
  <tr>
    <td align='center' valign='middle'>
		<table border="0" cellspacing="0" cellpadding="0" height="265">
			<tr>
                    <td colspan="4" align="center" height="20">
                        <p class="fontTD">Administração do portal da <%= iTitulo %></p>
                    </td>
				</tr>
			<tr>
				<td colspan="4" height="24" valign="bottom"><img src="images/log_01.jpg" width="540"></td>
				</tr>
                <tr>
				<td width="20" background="images/log_02.jpg" height="150">&nbsp;</td>
				<td width="250" height="150" bgcolor="#FFFFFF"><img src="images/logo.jpg" width="250" height="160"></td>
                    <td width="250" bgcolor="white" align="center" height="150">
                        <TABLE cellSpacing=0 cellPadding=5 align=center border=0>
					 <form action="pm_verifica.asp" method="post" onsubmit="return Validator(this)">
                                <TR>
                                    <td align="center" colspan="2">
<%
mensagem = Request("mes")
If mensagem = "1" Then
%>
                                        <p class="fontTD"><font color="yellow"><span style="background-color:red;">&nbsp;Usuário ou senha não conferem!&nbsp;</span></font></p>
<% ElseIf mensagem = "2" Then %>
                                        <p class="fontTD"><font color="yellow"><span style="background-color:red;">&nbsp;Entrada não autorizada!&nbsp;</span></font></p>
<% ElseIf mensagem = "3" Then %>
                                        <p class="fontTD"><span style="background-color:red;"><font color="yellow">&nbsp;Obrigado por usar nossos serviços!&nbsp;</font></span></p>
										
<% End If %>
                                    </td>
                                </TR>
                            <TR>
                                <TD><FONT face="Arial, Helvetica, sans-serif" size=2>Usuário</FONT></TD>
                                <TD>
									<INPUT size=22 name=username>
                                </TD>
                            </TR>
                            <TR>
                                <TD><FONT face="Arial, Helvetica, sans-serif" size=2>Senha</FONT></TD>
                                <TD>
									<INPUT type=password size=22 name=password> 
                                </TD>
                            </TR>
                                <TR>
                                    <td colspan="2" align="center" bgcolor="#E4E6E6">
                                        <p><INPUT type=image height=23 alt=Login width=60 src="images/button-login.gif" border=0></td>
                                </TR>
				  </FORM>
                        </TABLE>
                    </td>
				<td width="20" background="images/log_03.jpg" height="150">&nbsp;</td>
                </tr>
			<tr>
				<td colspan="4" height="22"><img src="images/log_04.jpg" width="540" height="30"></td>
				</tr>
			<tr>
                    <td colspan="4" height="18" align="center">
                        <p class="fontTD">CopyRight© GNove WebStudio - Todos os direitos reservados.</p>
                    </td>
				</tr>
		</table>

	</td>
  </tr>
</table>
</body>
</html>







