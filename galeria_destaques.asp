<%
randomize
set result = objConect.Execute("Select * From GALERIA_FOTOS")
nums = ""
while not result.eof
 nums = nums & result("a00idg") & "|"
 result.movenext
wend

arrTemp = Split(nums,"|")
id_sort = Round((ubound(arrTemp)-1) * Rnd)

sql_dest = "SELECT * FROM GALERIA_FOTOS Where a00idg = " & arrTemp(id_sort)

    Set result = Server.CreateObject("ADODB.Recordset")
    result.Open sql_dest, objConect, 3, 3

	SFoto = result(2)
	sNoticia = result(0)
	if sNoticia <> "" Then
%>
<table border="0">
    <tr>
        <td align="center" valign="middle">
		<a href="#" onClick="ajax('galeria.asp','conteudo')"><img src="config/galeria/<%= sFoto %>" alt="" width="103" height="77" border="0"></a>
		</td>
        <td height="67">
            <span class="navtext">Conheça nossa galeria de fotos<br><a href="#" onClick="ajax('galeria.asp','conteudo')"><%= result(4) %></span></a>
        </td>
    </tr>
</table>

<%
	result.Close
	Set result = Nothing
	End If
%>