<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">HISTÓRIA DO MUNICÍPIO DE <%= iCidade %></font></td>
							</tr>
            <table border="0" width="98%" align="center" cellspacing="0" bordercolordark="white" bordercolorlight="white">
               <tr>
                <td class="navtext">
<%
SQL = "SELECT * from historia Where a00id=1"
Set result = objConect.Execute(SQL)
%>
<%= result(2) %>
<%
result.close
Set result = Nothing
%>
                 </td>
                </tr>
             </table>