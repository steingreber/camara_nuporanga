<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
    <tr>
        <td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext" colspan="3">&nbsp;<font color="<%= ca05corLink %>">NOTÍCIA EM DESTAQUE</font></td>
    </tr>
    <tr>
        <td height="20" colspan="3" align="center" class="bodytext">
<!--#include file="noticiasdestaque.asp"-->
        </td>
    </tr>
    <tr>
        <td height="20" class="tittext" colspan="3">&nbsp;</td>
    </tr>
    <tr>
        <td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext" colspan="3">&nbsp;<font color="<%= ca05corLink %>">ÚLTIMAS NOTÍCIAS</font></td>
    </tr>
<%
Dim SQL1
SQL1 = "Select Top " & iNoticias & " * From NOTICIAS Where a06destaque='S' Order By a01data Desc"
    Set rsTempG = Server.CreateObject("ADODB.Recordset")
    rsTempG.Open SQL1, objConect, 3, 3
	Do While Not rsTempG.EOF
%>
    <tr>
        <td height="20" colspan="3" width="1054" class="bodytext">
            <p><strong>»</strong>&nbsp;<a href="#" onClick="ajax('noticias_view.asp?id=<%= rsTempG(0) %>','conteudo')"><%= Left(UCase(rsTempG(2)),56) %>...</a></p>
        </td>
    </tr>
<%
rsTempG.MoveNext
Loop
rsTempG.Close
Set rsTempG = Nothing

SQL11 = "Select * From MENUS where a00id=22"
    Set rsTempM = Server.CreateObject("ADODB.Recordset")
    rsTempM.Open SQL11, objConect, 3, 3
If rsTempM(4)="1" Then
%>
    <tr>
        <td height="20" class="tittext" colspan="3">&nbsp;</td>
    </tr>
							<tr>
        <td height="20" class="tittext" bgcolor="<%= cCorTopoBase %>" width="200">&nbsp;<font color="<%= ca05corLink %>">CONHEÇA NOSSO TURISMO</font></td>
        <td width="7" height="20" class="tittext">&nbsp;</td>
        <td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext">&nbsp;<font color="<%= ca05corLink %>">GALERIA DE FOTOS</font></td>
							</tr>
							<tr>
        <td height="20" align="center" class="tittext"><!--#include file="turismo_destaques.asp"--></td>
        <td width="7" height="20" class="tittext">&nbsp;</td>
        <td height="20" align="center" class="tittext"><!--#include file="galeria_destaques.asp"--></td>
							</tr>
<%
End If
rsTempM.Close
Set rsTempM = Nothing
%>
    </table>