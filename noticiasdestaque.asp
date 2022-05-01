<%
randomize
set result = objConect.Execute("Select * From NOTICIAS Where a06destaque='S' AND a05img<>'' Order By a01data Desc")
nums = ""
while not result.eof
 nums = nums & result("a00id") & "|"
 result.movenext
wend
If result.BOF = False Then
arrTemp = Split(nums,"|")
id_sort = Round((ubound(arrTemp)-1) * Rnd)

sql_dest = "SELECT * FROM NOTICIAS Where a00id = " & arrTemp(id_sort)

    Set result = Server.CreateObject("ADODB.Recordset")
    result.Open sql_dest, objConect, 3, 3


	Dim sFoto, sNoticia
	SFoto = result(5)
	sNoticia = result(0)
	if sNoticia <> "" Then
%>
<table border="0" height="170">
    <tr>
        <td width="230" rowspan="2" height="172" align="center" valign="middle">
		<IMG src="config/noticias/<%= sFoto %>" Id="imagem" OnLoad="if (document.getElementById('imagem').width > 200) imagem.width=200; if (document.getElementById('imagem').height > 135) imagem.height=135">
		</td>
        <td width="289" height="18">&nbsp;</td>
    </tr>
    <tr>
        <td width="289" height="150" valign="top">
		    <p align="justify" style="margin-left: 5pt; margin-right: 5px;">
			<span class="navtext"><strong><%= result(2) %></strong></span>
			<br>
			<span class="navtext"><a href="#" onClick="ajax('noticias_view.asp?id=<%= result(0) %>','conteudo')"><%= Trim(Left(result(3),100)) %>...</span></a>
			</p>
		</td>
    </tr>
</table>

<%
	result.Close
	Set result = Nothing
	End If
	End If
%>