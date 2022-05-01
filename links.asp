<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">LINKS ÚTEIS</font></td>
							</tr>
<%
SQL = "SELECT * from links order by a00id Desc"
Set result = objConect.Execute(SQL)
Do While Not result.EOF
%>							<tr>
							    <td height="20"><img src="imagens\marcador.gif" width="14" height="19" border="0"></td>
                                <td height="25" valign="top" class="navtext"><a href="<%= result("a02link") %>" target="_blank"><%= result("a01texto") %></a></td>
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