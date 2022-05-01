<%
On error resume next
randomize
set result = objConect.Execute("Select * From TURISMO")
nums = ""
while not result.eof
 nums = nums & result("A00IDTUR") & "|"
 result.movenext
wend

arrTemp = Split(nums,"|")
id_sort = Round((ubound(arrTemp)-1) * Rnd)

sql_dest = "SELECT * FROM TURISMO Where A00IDTUR = " & arrTemp(id_sort)

    Set result = Server.CreateObject("ADODB.Recordset")
    result.Open sql_dest, objConect, 3, 3

	SFoto = result(8)
	sNoticia = result(0)
	if sNoticia <> "" Then
%>
<table border="0">
    <tr>
        <td align="center" valign="middle">
		<a href="#" onClick="ajax('turismo.asp','conteudo')"><img src="config/turismo/<%= sFoto %>" alt="" width="103" height="77" border="0"></a>
		</td>
        <td height="67">
            <span class="navtext"><a href="#" onClick="ajax('turimo.asp','conteudo')"><%= result(2) %></span></a>
        </td>
    </tr>
</table>

<%
	result.Close
	Set result = Nothing
	End If
%>