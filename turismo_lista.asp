<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">TURISMO</font></td>
							</tr>
<%
SQL = "Select * From turismo Where A01CATEG =" & Request("id") & " Order By A02NOME"
Set result = objConect.Execute(SQL)
Do While Not result.EOF
%>							<tr>
                                <td width="533">
								<span class="navtext">
								<a href="#" onClick="ajax('turismo_view.asp?id=<%= result("A00IDTUR") %>','conteudo')"><%= result("A02NOME") %></a><br>
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
    <tr>
        <td align="center"><br><a href="#" onClick="ajax('turismo.asp','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a></td>
    </tr>
                        </table>