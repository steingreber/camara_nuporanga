            <table width="155" cellpadding="0" cellspacing="0">
                <tr>
                    <td height="4" colspan="3"></td>
                </tr>
<% If iBannerP = "1" Then %>
                <tr>
                    <td width="16" height="19"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a3.gif" width="18" height="21" border="0"></td>
                    <td width="114" height="19" align="center" bgcolor="<%= ca02corTopoTabela %>"><b><font size="2" face="Arial" color="white">Publicidade</font></b></td>
                    <td width="18" height="19"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a4.gif" width="18" height="21" border="0"></td>
                </tr>
                <tr>
                    <td height="5" colspan="3"></td>
                </tr>
                <tr>
                    <td width="16" height="10" align="right" valign="bottom"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a8.gif" width="18" height="10" border="0"></td>
                    <td width="114" height="10" align="right" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="10" valign="bottom"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a7.gif" width="17" height="10" border="0"></td>
                </tr>
                <tr>
                    <td height="59" colspan="3" bgcolor="<%= ca03corCorpoTabela %>" align="center">
<!--#include file="bannerP.asp"-->
					</td>
                </tr>
                <tr>
                    <td width="16" height="15"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a6.gif" width="18" height="20" border="0"></td>
                    <td width="114" height="15" align="center" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="15" align="left" valign="top"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a5.gif" width="17" height="18" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="8" colspan="3"></td>
                </tr>
<% End If %>
                <tr>
                    <td width="16" height="14"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a3.gif" width="18" height="21" border="0"></td>
                    <td width="114" height="14" align="center" bgcolor="<%= ca02corTopoTabela %>"><b><font size="2" face="Arial" color="white">Utilidades</font></b></td>
                    <td width="18" height="14" align="left"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a4.gif" width="18" height="21" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="5" colspan="3"></td>
                </tr>

                <tr>
                    <td width="16" height="6"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a8.gif" width="18" height="10" border="0"></td>
                    <td width="114" height="6" align="center" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="6" align="left"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a7.gif" width="17" height="10" border="0"></td>
                </tr>
<%
SQL = "SELECT * from utilidades order by a01texto"
Set result = objConect.Execute(SQL)
Do While Not result.EOF
%>	
                <tr>
                    <td width="16" height="12" bgcolor="<%= ca03corCorpoTabela %>">&nbsp;</td>
                    <td width="132" height="12" bgcolor="<%= ca03corCorpoTabela %>" colspan="2" onMouseOut="mOut(this,'<%=ca03corCorpoTabela%>');" onMouseOver="mOvr(this,'#f5f5f5');" onClick="mClk(this);"><a href="<%= result(2) %>" target="_blank"><font face="Arial" size="1" color="#000000"><%= UCase(result(1)) %></font></a></td>
                </tr>
                <tr>
                    <td width="16" height="1" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="114" height="1" align="center" bgcolor="<%= ca03corCorpoTabela %>" background="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/pontilhado.gif"></td>
                    <td width="18" height="1" align="center" bgcolor="<%= ca03corCorpoTabela %>" background="do.gif"></td>
                </tr>
<%
result.movenext
Loop
result.close
Set result = Nothing
%>
                <tr>
                    <td width="16" height="7"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a6.gif" width="18" height="20" border="0"></td>
                    <td width="114" height="7" align="center" bgcolor="<%= ca03corCorpoTabela %>">&nbsp;</td>
                    <td width="18" height="7" align="left" valign="top"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a5.gif" width="17" height="18" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="8" colspan="3"></td>
                </tr>
<% If enquete = "1" Then %>
                <tr>
                    <td width="16" height="14"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a3.gif" width="18" height="21" border="0"></td>
                    <td width="114" height="14" align="center" bgcolor="<%= ca02corTopoTabela %>"><b><font size="2" face="Arial" color="white">Enquete</font></b></td>
                    <td width="18" height="14" align="left"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a4.gif" width="18" height="21" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="5" colspan="3"></td>
                </tr>
                <tr>
                    <td width="16" height="6"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a8.gif" width="18" height="10" border="0"></td>
                    <td width="114" height="6" align="center" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="6" align="left"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a7.gif" width="17" height="10" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="12" align="center" bgcolor="<%= ca03corCorpoTabela %>" colspan="3">
<%
	sqly = "Select * From enquetes Where Active = True"
	Set objRsy = Server.CreateObject("ADODB.Recordset")
	objRsy.Open sqly, objConect, 3, 3
	objRsy.MoveLast
%>
			<strong><a href="javascript:open_window('enquete', 'enquete.asp?eId=<%= Trim(objRsy(0)) %>', 200, 40, 314, 451, 0, 0, 0, 0, 0)"><font class="navtext"><%= Trim(objRsy(2)) %></font></a></strong>
<%objRsy.close%>
                    </td>
                </tr>
                <tr>
                    <td width="16" height="1" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="1" align="center" bgcolor="<%= ca03corCorpoTabela %>" background="do.gif"></td>
                </tr>
                <tr>
                    <td width="16" height="7"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a6.gif" width="18" height="20" border="0"></td>
                    <td width="114" height="7" align="center" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="7" align="left" valign="top"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a5.gif" width="17" height="18" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="8" colspan="3"></td>
                </tr>
<% End If %>				
<% If clima = "1" Then %>
                <tr>
                    <td width="16" height="14"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a3.gif" width="18" height="21" border="0"></td>
                    <td width="114" height="14" align="center" bgcolor="<%= ca02corTopoTabela %>"><b><font size="2" face="Arial" color="white">Tempo</font></b></td>
                    <td width="18" height="14" align="left"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a4.gif" width="18" height="21" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="5" colspan="3"></td>
                </tr>
                <tr>
                    <td width="16" height="6"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a8.gif" width="18" height="10" border="0"></td>
                    <td width="114" height="6" align="center" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="6" align="left"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a7.gif" width="17" height="10" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="12" align="center" bgcolor="<%= ca03corCorpoTabela %>" colspan="3">
<%
SQL = "SELECT * from CLIMA Where a00id=1"
Set result = objConect.Execute(SQL)
Response.Write result(1)
result.Close
%>
                    </td>
                </tr>
                <tr>
                    <td width="16" height="1" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="1" align="center" bgcolor="<%= ca03corCorpoTabela %>" background="do.gif"></td>
                </tr>
                <tr>
                    <td width="16" height="7"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a6.gif" width="18" height="20" border="0"></td>
                    <td width="114" height="7" align="center" bgcolor="<%= ca03corCorpoTabela %>"></td>
                    <td width="18" height="7" valign="top"><img src="imagens/tamplate_<%= iTemplate %>/color_<%= iCores %>/a5.gif" width="17" height="18" border="0"></td>
                </tr>
                <tr>
                    <td width="150" height="7" colspan="3">&nbsp;</td>
                </tr>
<% End If %>
            </table>