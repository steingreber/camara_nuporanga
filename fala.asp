<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td colspan="2" height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">FALA PREFEITO</font></td>
							</tr>
<%
SQL = "Select * From fala"
Set RS = objConect.Execute(SQL)
%>
                <tr>
                    <td width="144" height="125" valign="top" align="center">
                        <p><img src="config/galeria/<%= RS(2) %>" width="133" height="200" border="0"><br>
						<span class="navtext"><%= RS(3) %></span></p>
                    </td>
					<td valign="top">
					<span class="navtext"><%= RS(1) %></span>
					</td>
                </tr>
<%
RS.Close
%>
            </table>