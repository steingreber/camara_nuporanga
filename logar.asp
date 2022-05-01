<%Response.ContentType = "text/html; charset=iso-8859-1"%>

<script language="JavaScript" src="suport/suport.js"></script>
<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">LOGAR NO PORTAL</font></td>
							</tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
    <tr>
        <td bgcolor="<%= ca02corTopoTabela %>">
            <p><span class="tittext">Entre com seus dados de acesso:</span></p>
        </td>
    </tr>
    <tr>
        <td height="70">
            <table width="100%" cellspacing="0" bordercolordark="white" bordercolorlight="black" cellpadding="0">
			<form action="logar_verifica.asp" method="post" name="FrmLogar" id="FrmLogar" onSubmit="return validalogin(this)">
                    <tr>
                    <td align="right" bgcolor="<%= ca03corCorpoTabela %>"><span class="navtext"><b>Email de acesso:</b></span></td>
                    <td bgcolor="<%= ca03corCorpoTabela %>"><input type=text name="email" size=20 class="ud_caixa"></td>
                        <td height="46" rowspan="2" align="center" valign="middle" bgcolor="<%= ca03corCorpoTabela %>">
                            <p><img align="absmiddle" src="imagens/bullet_chave.gif" width="21" height="13" border="0"><span class="navtext"><b>Não se lembra da senha? <a href="javascript:open_window('lembrar', 'lembrar_senha.asp', 230, 215, 308, 248, 0, 0, 0, 0, 0)"><br>Clique aqui!</a>.</b></span></p>
							<p><a href="#" onClick="javascript:ajax('assine_ja.asp','conteudo')"><span class="navtext">Não sou cadastrado!</span></a></p>
                        </td>
                    </tr>
                <tr>
                    <td align="right" bgcolor="<%= ca03corCorpoTabela %>"><span class="navtext"><b>Senha:</b></span></td>
                    <td bgcolor="<%= ca03corCorpoTabela %>">
                        <p><input type="password" name="senha" size="20" class="ud_caixa"></td>
                </tr>
                <tr>
                    <td colspan="2" height="39" align="center"><input type="image" name="formimage2" src="imagens/botao_login.gif" align="absmiddle" width="88" height="16" border="0"></td>
                    <td height="41" align="right" valign="middle">&nbsp;</td>
                </tr>
            </form>
            </table>
        </td>
    </tr>
</table>