<table width="151" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="17" align="center" bgcolor="#0000CC">&nbsp;<font face="Verdana" size="1" color="white"><B>ADMINSTRAÇÃO DO PORTAL</B></font></td>
	</tr>
	<tr>
		<td bgcolor="<%=cortop%>"></td>
	</tr>
	<tr>
		<td bgcolor="<%=cordown%>"></td>
	</tr>
	<tr>
		<td class="ms_href" bgcolor="#ECECEC" height=17 onMouseOut="mOut(this,'#ECECEC');" onMouseOver="mOvr(this,'entrada.asp');" onClick="mClk(this);"><a href="entrada.asp" class="menu"><img src="images/fundo_lado.gif" alt="" border="0" align="absmiddle">&nbsp;<font face="Verdana" size="1">INÍCIO</font></a></td>
	</tr>
<%
If Session("nome")="GNOVE WEB STUDIO" Then
SQL = "Select * From MENU_ADM Order By A01MENU"
Else
SQL = "Select * From MENU_ADM Where A03EXIBIR = 1 Order By A01MENU"
End If
Set RS = objConect.Execute(SQL)
Do While Not RS.EOF
If RS(3) = "0" Then
iCor = "#cccccc"
Else
iCor = "#ECECEC"
End If
%>

	<tr>
		<td class="ms_href" bgcolor="<%= iCor %>" height=15 onMouseOut="mOut(this,'<%= iCor %>');" onMouseOver="mOvr(this,'<%=cormOvr%>');" onClick="mClk(this);"><a href="<%= RS(2) %>" <% If RS(3) = "_blank" Then %>target="_blank"<% End If %> class="menu"><img src="images/fundo_lado.gif" alt="" width="11" height="15" border="0" align="absmiddle">&nbsp;<font face="Verdana" size="1"><%= RS(1) %></font></a></td>
	</tr>
	<tr>
		<td bgcolor="<%=cortop%>"></td>
	</tr>
<%
RS.MoveNext
Loop
RS.Close
%>
	<tr>
		<td bgcolor="<%=cortop%>"></td>
	</tr>
	<tr>
		<td bgcolor="<%=cordown%>"></td>
	</tr>
	<tr>
		<td bgcolor="#0000CC" align="center" valign="top">&nbsp;</td>
	</tr>
</table>
