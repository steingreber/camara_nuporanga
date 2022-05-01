<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0" bordercolordark="white" height="267">
							<tr>
							<td height="20" colspan="4" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">TURISMO</font></td>
							</tr>
<%
SQL = "Select * From turismo Where A00IDTUR =" & Request("id")
Set result = objConect.Execute(SQL)
%>							<tr>
                                <td height="116" valign="top" colspan="4">
								<span class="titulotext""><%= result("A02NOME") %></span><br>
								<span class="navtext"><%= result("A03END") & "<br>Fone: " & result("A04FONE") %><br>
								<%= "<strong>E-Mail:</strong> <a href='mailto:" & result("A05EMAIL") & "'>" & result("A05EMAIL") & "</a><br>" %>
								<%= "<strong>Web:</strong> <a href='" & result("A06WEB") & "' target='_blank'>" & result("A06WEB") & "</a><br><br>" %>
								<%= "<strong>Mais informações:</strong> " & result("A07DESC") & "<br>" %>
								</span>
                                </td>
                            </tr>
								<% If result("A08IMG01") <> "" Then %>
							<tr>
                                <td width="147" height="125" align="center">
								<a href="javascript:fotoZoom('config/turismo/<%= result("A08IMG01") %>')"><img src="config/turismo/<%= result("A08IMG01") %>" alt="" width="100" height="100" border="0"></a>
								</td>
                                <td width="147" height="125" align="center">
								<% If result("A09IMG02") <> "" Then %>
								<a href="javascript:fotoZoom('config/turismo/<%= result("A09IMG02") %>')"><img src="config/turismo/<%= result("A09IMG02") %>" alt="" width="100" height="100" border="0"></a>
								<% End If %>
								</td>
                                <td width="147" height="125" align="center">
								<% If result("A10IMG03") <> "" Then %>
								<a href="javascript:fotoZoom('config/turismo/<%= result("A10IMG03") %>')"><img src="config/turismo/<%= result("A10IMG03") %>" alt="" width="100" height="100" border="0"></a>
								<% End If %>
								</td>
                                <td width="147" height="125" align="center">
								<% If result("A11IMG04") <> "" Then %>
								<a href="javascript:fotoZoom('config/turismo/<%= result("A11IMG04") %>')"><img src="config/turismo/<%= result("A11IMG04") %>" alt="" width="100" height="100" border="0"></a>
								<% End If %>
								</td>
                            </tr>
								<% End If %>
<%
result.close
Set result = Nothing
%>
								<tr>
								<td><a href="#" onClick="ajax('turismo_lista.asp?id=<%= Request("id") %>','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a></td>
								</tr>
                        </table>
