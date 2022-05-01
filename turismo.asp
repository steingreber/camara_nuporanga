<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">TURISMO</font></td>
							</tr>
<%
SQL = "SELECT * from turismo_categ order by A01CATEG"
Set result = objConect.Execute(SQL)
Do While Not result.EOF
%>							<tr>
							    <td height="20"><img src="imagens\marcador.gif" width="14" height="19" border="0"></td>
                                <td height="25" valign="top" class="navtext"><a href="#" onClick="ajax('turismo_lista.asp?id=<%= result(0) %>','conteudo')"><%= result(1) %></a></td>
                            </tr>
<%
result.movenext
Loop
result.close
Set result = Nothing
%>
                        </table>