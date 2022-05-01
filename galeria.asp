<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20"  colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">GALERIA DE FOTOS DO MUNÍCIPIO DE <%= iCidade %> (clique na foto)</font></td>
							</tr>
<%
	sql = "SELECT * FROM Galeria Order By a00id desc"
    Set rsTemp = Server.CreateObject("ADODB.Recordset")
    rsTemp.Open sql, objConect, 3, 3
	Do While not rsTemp.EOF
%>
                <tr>
                    <td width="133" height="96" align="center"><a href="#" onclick="javascript:ajax('galeria_fotos.asp?idgaleria=<%= rsTemp(0) %>','conteudo','carregando...')"><img src="config/galeria/<%= rsTemp(4) %>" width="103" height="77" border="0"></a></td>
					<td width="330" height="96" class="navtext">
					<strong>Ocasião: </strong><%= rsTemp(3) %><br>
					<strong>Local: </strong><%= rsTemp(1) %><br>
					<strong>Data: </strong><%= rsTemp(2) %>
					</td>
                </tr>
<%
	rsTemp.MoveNext
	Loop
%>
            </table>
        </td>
    </tr>
</table>