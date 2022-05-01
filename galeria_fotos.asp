<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">GALERIA DE FOTOS DO MUNÍCIPIO DE <%= iCidade %> - (clique na foto)</font></td>
							</tr>
						</table>
                        <table border="0" width="99%" align="center" cellspacing="0">
<%
	sFoto = 0
	contador = 1
	sID   = Request("idgaleria")
	sql = "SELECT * FROM Galeria_fotos Where a01galeria=" & sID
    Set rsTemp = Server.CreateObject("ADODB.Recordset")
    rsTemp.Open sql, objConect, 3, 3
	Do While not rsTemp.EOF
%>
                    <td align="center"><a href="javascript:open_window('win', 'palco.asp?Page=<%= contador %>&galeria=<%= rsTemp(1) %>', 50, 50, 601, 311, 0, 0, 0, 0, 0);" target="_self"><img src="config/galeria/<%= rsTemp(2) %>" width="103" height="77" border="0"></a></td>
<%
	sFoto = sFoto + 1
	contador = contador+1
If sFoto = 4 then
	Response.Write "            </tr>" & vbCrLf
	Response.Write "            <tr>"
	sFoto = 0
End If

	rsTemp.MoveNext
	Loop
	rsTemp.close
%>
            </table>
<p align="center"><a href="#" onclick="javascript:ajax('galeria.asp','conteudo','carregando...')"><img src="imagens/botao_voltar.gif" alt="" border="0"></a></p>
        </td>
    </tr>
</table>
<br><br>