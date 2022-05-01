<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">ADMINISTRAÇÃO DO MUNICÍPIO</font></td>
							</tr>

<%
SQL = "Select * From administracao"
Set RS = objConect.Execute(SQL)
%>
                            <tr>
                                <td width="602" class="navtext">
                                    <p align="justify">
                                    <% Response.Write RS(1) %>
                                    </p>
                                </td>
                            </tr>
<%
RS.Close
Set RS = Nothing
%>
             </table>