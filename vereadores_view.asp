<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">VEREADORES</font></td>
							</tr>
<%
Dim aId
aId = Request("id")
SQL = "SELECT * from vereadores Where a00id = " & aId
Set result = objConect.Execute(SQL)
%>
                <tr>
                    <td width="311" align="center" height="209">
                        <p class="navtext"><%= result(1) %></p>
                        <p class="navtext"><%= result(3) %></p>
                    </td>
                    <td width="311" height="209" align="center"><img align="absmiddle" src="config/vereadores/<%= result(5) %>" border="0"></td>
                </tr>
                <tr>
                    <td height="42" colspan="2">
                        <p align="justify"><span class="navtext"><%= result(4) %></span></p>
                    </td>
                </tr>
                <tr>
                    <td height="42" colspan="2" align="center"><span class="navtext"><a href="mailto:<%= LCase(result(2)) %>"><font color="#BB000D"><b>Mandar email para este(a) vereador(a)</b></font></a></span></td>
                </tr>
                <tr>
                    <td height="42" align="center" colspan="2"><a href="#" onClick="ajax('vereadores.asp','conteudo')"><img src="imagens/botao_anterior.gif" alt="" border="0"></a></td>
                </tr>
<%
result.close
Set result = Nothing
%>
            </table>