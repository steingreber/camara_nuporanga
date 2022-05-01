<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="3" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">VEREADORES</font></td>
							</tr>
                <tr>
                    <td align="center" colspan="3">
                        <p><FONT color=#0000ff size=2><span class="navtext"><%= legislatura %></span></FONT></p>
                    </td>
                </tr>
                <tr>
                    <td colspan="3" height="19" align="center">
                        <p><span class="navtext"><font size="1" color="#BB000D">(clique no nome do vereador para acessar seus dados)</font></span></p>
                    </td>
                </tr>
                <tr>
                    <td bgcolor="#E8E8E8" class="navtext">&nbsp;Nome</td>
                    <td bgcolor="#E8E8E8" class="navtext">&nbsp;Email</td>
                    <td bgcolor="#E8E8E8" class="navtext">&nbsp;Partido</td>
                </tr>
<%
SQL = "SELECT * from vereadores order by a01nome"
Set result = objConect.Execute(SQL)
Do While Not result.EOF
%>
                <tr>
                    <td class="navtext"><a href="#" onClick="ajax('vereadores_view.asp?id=<%= result(0) %>','conteudo')"><%= UCase(result(1)) %></a></td>
                    <td class="navtext"><a href="mailto:<%= LCase(result(2)) %>"><%= LCase(result(2)) %></a></td>
                    <td class="navtext"><%= UCase(result(3)) %></td>
                </tr>
<%
result.movenext
Loop
result.close
Set result = Nothing
%>
            </table>