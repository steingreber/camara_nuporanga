<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">SECRETARIAS</font></td>
							</tr>
<%
Dim aId
aId = Request("id")
SQL = "SELECT * from secretarias Where a00id = " & aId
Set result = objConect.Execute(SQL)
%>
                <tr>
                    <td width="475" height="81" valign="top">
                        <p><span class="tittext"><strong><%= result(1) %></strong></span></p>
                        <p align="justify"><span class="navtext"><%= result(2) %></span></p>
                    </td>
<%
result.close
Set result = Nothing
%>
                </tr>
                <tr>
                    <td width="475" height="55" align="center"><a href="#" onClick="javascript:ajax('secretarias.asp','conteudo')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a></td>
                </tr>
            </table>