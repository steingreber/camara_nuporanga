<%Response.ContentType = "text/html; charset=iso-8859-1"%>

<script language="JavaScript" src="suport/suport.js"></script>
<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">FAÇA SEU CADASTRO</font></td>
							</tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
    <tr>
        <td>
<%
Dim mensagem
mensagem = Request("msg")
If mensagem = "1" Then
	Response.Write "<p align='center'><span class='tittext'><strong><font size='3' color='#FF0000'>Para ter acesso a este material você precisa ser cadastrado e estar logado no portal!</font></strong></span></p>"
End If
%>
        </td>
    </tr>
    <tr>
        <td bgcolor="<%=ca02corTopoTabela%>">
            <p><span class="tittext">Já tenho cadastro:</span></p>
        </td>
    </tr>
    <tr>
        <td height="70">
            <table width="100%" cellspacing="0" bordercolordark="white" bordercolorlight="black" cellpadding="0">
			<form action="logar_verifica.asp" method="post" name="FrmLogar" id="FrmLogar" onSubmit="return validalogin(this)">
                    <tr>
                    <td align="right" bgcolor="<%= ca03corCorpoTabela %>"><span class="navtext"><b>Email de acesso:</b></span></td>
                    <td bgcolor="<%= ca03corCorpoTabela %>"><input type=text name="email" size=20 class="ud_caixa"></td>
                        <td rowspan="2" align="right" valign="middle" height="46" bgcolor="<%= ca03corCorpoTabela %>">
                            <p><img align="absmiddle" src="imagens/bullet_chave.gif" width="21" height="13" border="0"><span class="navtext"><b>Não se lembra da senha?<br><a href="javascript:open_window('lembrar', 'lembrar_senha.asp', 230, 215, 308, 248, 0, 0, 0, 0, 0)"><span style="background-color:yellow;"><font color="#FF0000">Clique aqui!</font></span></a>.</b></span></p>
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
    <tr>
        <td align="center">
            <p>&nbsp;</p>
        </td>
    </tr>
    <tr>
        <td bgcolor="<%=ca02corTopoTabela%>"><span class="tittext">Ainda não sou cadastrado:</span></td>
    </tr>
    <tr>
        <td>
<form action="assine_ok.asp" method="post" onSubmit="return validacadastro(this)">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="143" bgcolor="<%= ca03corCorpoTabela %>" align="right"><img align="absmiddle" src="imagens/bullet_003.gif" width="11" height="11" border="0"></td>
                    <td width="314" bgcolor="#777777">
                        <p><span class="tittext"><b><font color="white">&nbsp;Dados de login:</font></b></span></p>
                    </td>
                </tr>
                    <tr>
                    <td width="143" align="right" height="25"><span class="navtext"><b>* Seu email:</b></span></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="25">
                            <p><input type="text" name="fu30username" size="23" maxlength="50" class="ud_caixa"><br><span class="navtext" style="background-color:yellow;"><font color="#FF0000">Informe um email válido para a confirmação de cadastro!</font></span></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" align="right" height="25"><span class="navtext"><b>* Senha:</b></span></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="25">
                            <p><input type=password name=fu31senha  size="20" maxlength=10 class="ud_caixa"></p>
                        </td>
                    </tr>
                    <tr>
                        <td width="975" align="right" colspan="2" height="5"></td>
                    </tr>
                <tr>
                    <td width="143" align="right" bgcolor="<%= ca03corCorpoTabela %>"><img align="absmiddle" src="imagens/bullet_003.gif" width="11" height="11" border="0"></td>
                    <td width="314" bgcolor="#777777"><span class="tittext"><b><font color="white">&nbsp;Dados do assinante:</font></b></span></td>
                </tr>
                <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* Nome:</b></td>
                    <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
					<input type="text" name="fu01Nome" size="40" class="ud_caixa">
                    </td>
                </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* Fone contato:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type="text" name="fu04FoneContato" size="20" maxlength="12" class="ud_caixa" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* Estado civil:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><select name="fu05EstadoCivil" class="ud_caixa" style="width: 100px;">
								<option selected>Selecione</option>
                                <option value="SOLTEIRO">SOLTEIRO</option>
                                <option value="CASADO">CASADO</option>
                                <option value="VIÚVO">VIÚVO</option>
                                <option value="DIVORCIADO">DIVORCIADO</option>
								<option value="UNIAO ESTAVEL">UNIAO ESTAVEL</option>
						</select> <b><span class="navtext">* Sexo:</span></b><select name="fu06Sexo" class="ud_caixa" style="width: 100px;">
								<option selected>Selecione</option>
                                <option value="M">Masculino</option>
                                <option value="F">Feminino</option>
						</select></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* Profissão:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type=text name=fu07Profissao  size="40" maxlength=37 class="ud_caixa"></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* RG:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type=text name=fu07Identidade  size="23" maxlength=37 class="ud_caixa"></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* CPF nº:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type=text name=fu06CPF  size="20" maxlength=20 class="ud_caixa"></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>End. residencial:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type=text name=fu14EndRes  size="24" maxlength=60 class="ud_caixa"><b><span class="navtext">Número</span></b> <input type="text" name="fu15Numero" size="5" maxlength="5" class="ud_caixa" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"></p>
                        </td>
                    </tr>
                    <tr>
                        <td width="143" align="right" class="navtext" height="24">
                            <p><b>Complemento:</b></p>
                        </td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type=text name=fu16Complemento  size="24" maxlength=20 class="ud_caixa"></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>Bairro:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type=text name=fu17Bairro  size="20" maxlength=30 class="ud_caixa"></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* Cidade:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type=text name=fu19Cidade size="30" maxlength=40 class="ud_caixa"><b><span class="navtext">* Estado</span></b> 
							<select name="fu20Estado" class="ud_caixa">
 								<option selected>--</option>
                                <option value=AC>AC</option>
                                <option value=AL>AL</option>
                                <option value=AM>AM</option>
                                <option value=AP>AP</option>
                                <option value=BA>BA</option>
                                <option value=CE>CE</option>
                                <option value=DF>DF</option>
                                <option value=ES>ES</option>
                                <option value=GO>GO</option>
                                <option value=MA>MA</option>
                                <option value=MG>MG</option>
                                <option value=MS>MS</option>
                                <option value=MT>MT</option>
                                <option value=PA>PA</option>
                                <option value=PB>PB</option>
                                <option value=PE>PE</option>
                                <option value=PI>PI</option>
                                <option value=PR>PR</option>
                                <option value=RJ>RJ</option>
                                <option value=RN>RN</option>
                                <option value=RO>RO</option>
                                <option value=RR>RR</option>
                                <option value=RS>RS</option>
                                <option value=SC>SC</option>
                                <option value=SE>SE</option>
                                <option value=SP>SP</option>
                                <option value=TO>TO</option>
						</select></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* Cep:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type="text" name="fu18CEP" size="9" maxlength="9" class="ud_caixa" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>Celular:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type="text" name="fu21FoneRes" size="20" maxlength="12" class="ud_caixa" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"></p>
                        </td>
                    </tr>
                    <tr>
                    <td width="143" class="navtext" align="right" height="24"><b>* Data nacimento:</b></td>
                        <td width="314" bgcolor="<%= ca03corCorpoTabela %>" height="24">
                            <p><input type=text name=fu02Data_Nascto  size=10 maxlength=10 class="ud_caixa"> <span class="navtext">dd/mm/aaaa</span></p>
                        </td>
                    </tr>
                    <tr>
                        <td width="457" align="center" colspan="2" height="42" valign="bottom">
						<input type="image" name="formimage1" src="imagens/botao_cadastro.gif" align="absmiddle" border="0" width="88" height="16"></td>
                    </tr>
            </table>
            </form>
        </td>
    </tr>
</table>