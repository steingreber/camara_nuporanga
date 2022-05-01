<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<%
Dim SQL, sSub, sCat
sCat = Request("cat")
sSub = Request("sub")

SQL = "Select * From CATEGORIAS Where A00IDC1=" & sCat
    Set rsTempC = Server.CreateObject("ADODB.Recordset")
    rsTempC.Open SQL, objConect, 3, 3

SQL2 = "Select * From SUBCATEG Where A00IDSUB=" & sSub
    Set rsTempS= Server.CreateObject("ADODB.Recordset")
    rsTempS.Open SQL2, objConect, 3, 3

Response.Write "<br><p class=""navtext"">Você esta em:<br><font color='#000000'><a href='default.asp'>HOME</a> - <a href='#' onClick=" & chr(34) & "ajax('verifica.asp?id="& sCat &"','conteudo')" & chr(34) & ">" & rsTempC(1) & "</a> - " & UCase(rsTempS(2)) & "</font><hr size='1' noshade></p>"

SQL = "Select * From MATERIAL Where A11SUBCATEG=" & sSub
    Set rsTempH = Server.CreateObject("ADODB.Recordset")
    rsTempH.Open SQL, objConect, 3, 3
	Response.Write "<ul type='square'>"
    Do While Not rsTempH.EOF %>
        <li><a href="#" onClick="ajax('view_material.asp?id=<%= rsTempH(0) %>&cat=<%= sCat %>&sub=<%= sSub %>','conteudo')"><span class="navtext"><%= rsTempH(1) %></span></a></li>
<%	rsTempH.MoveNext
	Loop
	Response.Write "</ul>"
	rsTempH.Close
	Set rsTempH = Nothing %>