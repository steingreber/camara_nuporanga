<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<html>

<head>
<title><%= iCidade %></title>
<meta name="robots" content="none">
<link type="text/css" rel="stylesheet" href="suport/suport_def.css">
</head>
<%
On error resume next
Dim Connect_String
Dim Page_Size
Dim StartPage
Dim Current_Page
Dim Page_Count
Dim CssClass
Dim CellValue
Dim intRowCount
Dim SQL, sID, sGaleria
Dim i, strTemp, corTR, corData

	sID = Request("id")
	sGaleria = Request("galeria")

   SQL = "SELECT * FROM Galeria_fotos Where a01galeria=" & sGaleria & " Order By a00idg"

   Set rsTemp = Server.CreateObject("ADODB.Recordset")
   rsTemp.CursorLocation = 3
   If Request("Page") = "" Then
      Current_Page = 1
   else
     Current_Page = CInt(Request("Page"))
   end if

   Page_Size = "1"
   rsTemp.PageSize = Page_Size
   rsTemp.Open SQL, objConect

   Page_Count = rsTemp.PageCount
   If Current_Page < 1 Then Current_Page = 1
   If Current_Page > Page_Count Then Current_Page = Page_Count
   If Page_Count > 0 then rsTemp.AbsolutePage = Current_Page
   Errou
%>
<body bgcolor="<%= cCorTopoBase %>" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="600" height="310" cellpadding="0" cellspacing="0">
<%
   Do While rsTemp.AbsolutePage = Current_Page And Not rsTemp.EOF
   If rsTemp.BOF AND rsTemp.EOF then response.End
%>
    <tr>
        <td width="411" rowspan="2" align="center"><img src="config/galeria/<%= rsTemp(3) %>" width="400" height="300" border="0"></td>
        <td width="189" align="center" class="tittext">Fotos do município de <%= iCidade %><BR><BR><%= rsTemp(4) %></td>
    </tr>
    <tr>
	<td align="center">
<%
If Current_Page > 1 then
  Response.Write "<A HREF=""palco.asp?Page=" & (Current_Page - 1) & "&galeria=" & sGaleria & """><img src='imagens/botao_anterior.gif' alt='Anterior' border='0'> </A> " & vbCrLfelse
  Response.Write ""
End If
If Current_Page < Page_Count then
  Response.Write "<A HREF=""palco.asp?Page=" & (Current_Page + 1) & "&galeria=" & sGaleria & """><img src='imagens/botao_proximo.gif' alt='Próximo' border='0'> </A>"
Else
  Response.Write ""
End If
%>
	</td>
    </tr>
<%
   rsTemp.MoveNext
   Loop
   rsTemp.Close
   Set rsTemp = Nothing
%>
</table>
<%
Sub Errou
If err <> 0 Then
%>
<br>
<table cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
        <td align="center" valign="middle">
            <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" align="center" cellpadding="0" cellspacing="0" width="571" bordercolordark="black" bordercolorlight="black">
                <tr>
                    <td width="571" height="63" align="center">
                        <p style="margin-left:5pt;"><span class="titulo">Informação do sistema...</span></p>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="128" align="center">
           <span class="titulo">Erro nº <%=err.Number%> <br><%=err.Source%><br><%= err.Description%></span>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="38" align="center">
           <span class="titulo"><a href="javascript:close();">Fechar</a></span>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<br>
<%
Response.End
End If
End Sub
%>

</body>

</html>
