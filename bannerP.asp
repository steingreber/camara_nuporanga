<table width="130" cellpadding="0" cellspacing="0">
<%
SQLp = "SELECT * FROM BANNERS Where a05exibir=1 AND a07tamanho=0 Order By a00id Desc"
Set rsBan = Server.CreateObject("ADODB.Recordset")
rsBan.Open SQLp, objConect, 3, 3


Do While Not rsBan.EOF
pImg1 = rsBan(3)
pdesc = rsBan(1)
pLink = rsBan(2)
pTipG = rsBan(8)
pNovo = rsBan(6)
If pTipG = "0" Then
%>
    <tr>
        <td width="130" height="54">
		<a href="<%= pLink %>" <% If pNovo = "0" Then %>target="_blank"<% End If %>><img src="config/banners/<%= pImg1 %>" alt="<%= pdesc %>" width="130" height="54" border="0"></a>
		</td>
    </tr>
<% Else %>
    <tr>
        <td width="130" height="54">
                    <script>ConteudoFlash('config/banners/<%= pImg1 %>','54','130');</script>
		</td>
    </tr>
<% End If %>
    <tr>
        <td width="130" height="5"></td>
    </tr>
<%
rsBan.MoveNext
Loop
rsBan.close
%>
</table>