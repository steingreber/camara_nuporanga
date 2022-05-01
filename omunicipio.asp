<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">DADOS DO MUNICÍPIO DE <%= iCidade %></font></td>
							</tr>
               <tr>
                 <td>
<%
SQL = "SELECT * from dados Where a00id=1"
Set result = objConect.Execute(SQL)
%>
<span class="navtext"><%= result(2) %></span>
<%
result.close
Set result = Nothing
%>
                   </td>
                </tr>
             </table>