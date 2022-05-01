<table width="159" cellpadding="0" cellspacing="0">
                <tr>
                    <td height="4" colspan="3" width="153"></td>
                </tr>
                <tr>
                    <td width="18" height="19"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a3.gif" width="18" height="21" border="0"></td>
                    <td width="120" height="19" align="center" bgcolor="<%= ca02corTopoTabela %>"><b><font size="2" face="Arial" color="white">Procura</font></b></td>
                    <td width="15" height="19"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a4.gif" width="18" height="21" border="0"></td>
                </tr>
                <tr>
                    <td height="5" colspan="3" width="153"></td>
                </tr>
                <tr>
                    <td width="18" height="10" align="right" valign="bottom"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a8.gif" width="18" height="10" border="0"></td>
                    <td width="120" height="10" align="right" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="15" height="10" valign="bottom"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a7.gif" width="17" height="10" border="0"></td>
                </tr>
                <tr>
                    <td width="153" height="18" colspan="3" bgcolor="<%= ca03corCorpoTabela %>" align="center"><b><font size="1" face="Arial">Digite o que procura...</font></b></td>
                </tr>
                <tr>
                        <form action="busca_ok.asp">
                    <td width="153" height="18" colspan="3" bgcolor="<%= ca03corCorpoTabela %>" align="center">
                        <input type="text" name="search" value="<% =Request.QueryString("search") %>" size="15" class="ud_caixa"><input type="submit" name="formbutton1" value="OK" class="ud_caixa">
						<input type="hidden" name="mode" value="allwords">
                    </td>
                        </form>
                </tr>
                <tr>
                    <td width="18" height="15"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a6.gif" width="18" height="20" border="0"></td>
                    <td width="120" height="15" align="center" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="15" height="15" align="left" valign="top"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a5.gif" width="17" height="18" border="0"></td>
                </tr>
                <tr>
                    <td width="153" height="8" colspan="3"></td>
                </tr>
                <tr>
                    <td width="18" height="14"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a3.gif" width="18" height="21" border="0"></td>
                    <td width="120" height="14" align="center" bgcolor="<%= ca02corTopoTabela %>"><b><font size="2" face="Arial" color="white">Menu</font></b></td>
                    <td width="15" height="14" align="left"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a4.gif" width="18" height="21" border="0"></td>
                </tr>
                <tr>
                    <td width="153" height="5" colspan="3"></td>
                </tr>
                <tr>
                    <td width="18" height="6"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a8.gif" width="18" height="10" border="0"></td>
                    <td width="120" height="6" align="center" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="15" height="6" align="left"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a7.gif" width="17" height="10" border="0"></td>
                </tr>
<%
SQL = "Select * From MENUS Where a04exibir = 1 Order By a01texto"
Set RS = objConect.Execute(SQL)
Do While Not RS.EOF
%>
                <tr>
                    <td width="18" height="12" bgcolor="<%= ca03corCorpoTabela %>">&nbsp;</td>
                    <td width="135" height="12" bgcolor="<%= ca03corCorpoTabela %>" colspan="2" height=17 onMouseOut="mOut(this,'<%=ca03corCorpoTabela%>');" onMouseOver="mOvr(this,'#f5f5f5');" onClick="mClk(this);"><a href="<%= Replace(RS(2),Chr(34),"'",1) %>" <% If RS(3) = "_blank" Then %>target="_blank"<% End If %>><font size="1" face="Arial" color="black"><%= RS(1) %></font></a></td>
                </tr>
                <tr>
                    <td width="18" height="1" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="120" height="1" align="center" bgcolor="<%= ca03corCorpoTabela %>" background="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/pontilhado.gif"></td>
                    <td width="15" height="1" align="center" bgcolor="<%= ca03corCorpoTabela %>" background="do.gif"></td>
                </tr>
<%
RS.MoveNext
Loop
RS.Close
%>
                <tr>
                    <td width="18" height="17"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a6.gif" width="18" height="20" border="0"></td>
                    <td width="120" height="17" align="center" bgcolor="<%= ca03corCorpoTabela %>">&nbsp;</td>
                    <td width="15" height="17" align="left" valign="top"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a5.gif" width="17" height="18" border="0"></td>
                </tr>
            </table>