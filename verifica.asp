<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<%
Dim sID, SQL, SQL2
sID = Request("id")

SQL = "Select * From CATEGORIAS Where A00IDC1=" & sID
    Set rsTempG = Server.CreateObject("ADODB.Recordset")
    rsTempG.Open SQL, objConect, 3, 3

Response.Write "<br><p class=""navtext"">Você esta em:<br><font color='#000000'><a href='default.asp'>HOME</a> - " & rsTempG(1) & "</font><hr size='1' noshade></p>"

If rsTempG(3) = "1" Then
Response.Write "<p align='center' class=""navtext""><font color='#FF0000'>Atenção!!!</font><BR>Alguns dos relatórios podem estar em formato PDF. Caso você não possua o programa Adobe Acrobat Reader, <a href='http://www.adobe.com/br/products/acrobat/readstep2.html' target='_blank'><font color=""#ff0000"">clique aqui</font></a> para baixa-lo.</p><br>"
Response.Write "<p class=""navtext""><font color='#000000'>Selecione a categoria desejada.</font></p>"
SQL2 = "Select * From SUBCATEG Where A01CATEG =" & rsTempG(0) & " Order By A02TITULO"
    Set rsTempH = Server.CreateObject("ADODB.Recordset")
    rsTempH.Open SQL2, objConect, 3, 3
	Response.Write "<ul type='square'>"
	Do While Not rsTempH.EOF

	Response.Write "<li><a href='#' onClick=" & chr(34) & "ajax('view.asp?sub=" & rsTempH(0) & "&cat=" & rsTempH(1) & "','conteudo')" & chr(34) & "><span class=""navtext"">" & rsTempH(2) & "</span></a></li>"

	rsTempH.MoveNext
	Loop
	Response.Write "</ul>"
Else

End If

rsTempG.Close
rsTempH.Close
objConect.Close
Set objConect = Nothing
Set rsTempG   = Nothing
Set rsTempH   = Nothing
%>