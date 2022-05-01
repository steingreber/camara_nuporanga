            <table cellpadding="0" cellspacing="1" width="98%">
<%
SQL = "SELECT * from turismo_categ order by A01CATEG"
Set result = objConect.Execute(SQL)
%>
                <tr>
                    <td height="125" valign="top" align="center">
                        <p><select name="formselect1" size="1" OnChange="namosw_goto_byselect(this, 'parent')" class="ud_caixa">
                            <option selected>Selecione a categoria</option>
                            <%Do While Not result.EOF%>
                            <option value="turismo_lista.asp?id=<%= result(0) %>"><%= result(1) %></option>
							<%
							result.movenext
							Loop
							result.close
							Set result = Nothing
							%>
							</select></p>
                        <table border="0" width="100%" align="center" cellspacing="0" bordercolordark="white" bordercolorlight="#C0C0C0">
<%
SQL = "Select * From turismo Where A01CATEG =" & Request("id") & " Order By A02NOME"
Set result = objConect.Execute(SQL)
Do While Not result.EOF
%>							<tr>
                                <td width="533">
								<span class="navtext">
								- <a href="turismo_view.asp?id=<%= result("A00IDTUR") %>"><font color="#0000FF"><%= result("A02NOME") %></font></a><br>
                                    <%= result("A03END") & "<br>Fone: " & result("A04FONE") %>
								</span>
                                </td>
                            </tr>
                            <%
result.movenext
Loop
result.close
Set result = Nothing
%>
                        </table>
                    </td>
                </tr>
            </table>