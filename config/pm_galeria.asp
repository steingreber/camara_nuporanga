<!--#include file="_cnx.asp"-->
<!--#include file="../config.asp"-->
<%
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"

	SQLPass = "Select * From login Where a02email = '" & Session("usuario") & "'"
	Set objRS = objConect.Execute(SQLPass)
	If Mid(objRS("a05permissoes"),7,1) = "0" Then
		response.redirect "entrada.asp?mes=1"
	End If
%>
<html>
<head>
<title>Galeria - PageMaster 2.0</title>
<meta name="robots" content="none">
<link rel="stylesheet" href="pagemaster.css">
<link type="text/css" rel="stylesheet" href="../suport/suport.css">
<script language="JavaScript" src="../suport/suport.js"></script>
<script language=JavaScript>
<!--
function confirm_delete() {
  return confirm ("Voc� realmente deseja excluir este registro?")
}
function fotoZoom(img){
  foto1= new Image();
  foto1.src=(img);
  Controlla(img);
}
function Controlla(img){
  if((foto1.width!=0)&&(foto1.height!=0)){
    viewFoto(img);
  }
  else{
    funzione="Controlla('"+img+"')";
    intervallo=setTimeout(funzione,20);
  }
}
function viewFoto(img){
  largh=foto1.width+20;
  altez=foto1.height+20;
  stringa="width="+largh+",height="+altez;
  finestra=window.open(img,"",stringa);
}
// -->
</script>

<%
'On error resume next
Dim Connect_String
Dim Page_Size
Dim StartPage
Dim Current_Page
Dim Page_Count
Dim CssClass
Dim CellValue
Dim intRowCount
Dim SQL
Dim i, strTemp, corTR, corData

sCampo = Request.Form("q")
sTipo  = Request.Form("t")
sValor = Request.Form("totext")

If Trim(sValor) = "" Then
   SQL = "Select * From [Galeria] Order By [a00id] Desc"
Else
   If Trim(sTipo) = "Like" Then
   SQL = "Select * From [Galeria] Where " & sCampo & " " & sTipo & "'%" & sValor & "%'"
   Else
   SQL = "Select * From [Galeria] Where " & sCampo & " " & sTipo & "'" & sValor & "'"
   End If
End If

   Set rsTemp = Server.CreateObject("ADODB.Recordset")
   rsTemp.CursorLocation = 3
   If Request("Page") = "" Then
      Current_Page = 1
   else
     Current_Page = CInt(Request("Page"))
   end if

   Page_Size = "50"
   rsTemp.PageSize = Page_Size
   rsTemp.Open SQL, objConect

   Page_Count = rsTemp.PageCount
   If Current_Page < 1 Then Current_Page = 1
   If Current_Page > Page_Count Then Current_Page = Page_Count
   If Page_Count > 0 then rsTemp.AbsolutePage = Current_Page
   Errou
%>
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000" leftmargin=0 topmargin=0>
<table width="100%" cellpadding="0" cellspacing="0">
    <tr>
        <td width="153" align="center">
<!--#include file="menu_lateral.asp"-->
</td>
        <td align="center" valign="top" class="Cabecalho_Calendario">
<table cellpadding=0 cellspacing=0 width="98%" align=center>
    <tr>
        <td height="5" bgcolor="#0000ff" width="100%">
        </td>
    </tr>
<tr>
	<td class="tabelatitulo">
		&nbsp;<span class="aFonte">��&nbsp;GALERIA DE FOTOS&nbsp;��</span>
	</td>
    <tr>
                    <td width="912" height=20 class="tabelatitulo">
                        <p><a href="pm_Galeria_opt.asp?acao=1"><img src="images/bt_novo.gif" alt="" border="0" align="middle"></a>
<span class="aFonte"><B>
                        <%
  Response.Write "<A HREF=""pm_Galeria.asp?Page=1""><img src='images/bt_primeiro.gif' alt=Primeiro border=0> </A>"
if Current_Page > 1 then
  Response.Write "<A HREF=""pm_Galeria.asp?Page=" & (Current_Page - 1) & """><img src='images/bt_anterior.gif' alt=Anterior border=0> </A> " & vbCrLfelse
  Response.Write ""
end if
if Current_Page < Page_Count then
  Response.Write "<A HREF=""pm_Galeria.asp?Page=" & (Current_Page + 1) & """><img src='images/bt_proximo.gif' alt=Pr�ximo border=0> </A>"
else
  Response.Write ""
end if
  Response.Write "<A HREF=""pm_Galeria.asp?Page=" & Page_Count & """><img src='images/bt_ultimo.gif' alt=�ltimo border=0>   </A>"
%>
</B></FONT></p>
                    </td>
                </tr>
    <tr>
        <td height=5 width=853 class="tabelaitem">
        </td>
    </tr>
    <tr>
        <td width=853 align="center" valign="top">
            <table align="center" border="1" cellspacing="0" width="100%" bordercolordark="white" bordercolorlight="black">
                <tr>
                    <td>
                        <table align="center" cellpadding="0" cellspacing="1" width="100%">
                            <tr>
                                <td width="63" align="center" class="tabelatitulo">
                                    <p><span class="aFonte">OP��ES</span></p>
                                </td>
                                <td align="center" class="tabelatitulo">
                                    <p><span class="aFonte">NOME DA GALERIA</span></p>
                                </td>
                                <td align="center" class="tabelatitulo">
                                    <p><span class="aFonte">LOCAL</span></p>
                                </td>
                                <td align="center" class="tabelatitulo">
                                    <p><span class="aFonte">DATA</span></p>
                                </td>
                                <td align="center" class="tabelatitulo">
                                    <p><span class="aFonte">FOTO</span></p>
                                </td>
                            </tr>
<%
   Dim pmCorSel
   pmCorSel = "#E7E7E7"
   Do While rsTemp.AbsolutePage = Current_Page And Not rsTemp.EOF
   If rsTemp.BOF AND rsTemp.EOF then response.End
   If pmCorSel = "#FFFFFF" Then: pmCorSel = "#E7E7E7": Else pmCorSel = "#FFFFFF"
%>
                            <tr bgcolor="<%= pmCorSel %>">
                                <td width="100">
                                    <table cellpadding="0" cellspacing="0" width="100%" align="center">
                                        <tr>
                                            <td align="center">
                                                <p><a href="pm_Galeria_opt.asp?acao=2&sF=<%= rsTemp("a00id") %>"><img src="images/botao_editar.gif" alt="Editar registro" width="16" height="16" border="0"></a></p>
                                            </td>
                                            <td align="center">
                                                <p><a onclick="return confirm_delete()" href=pm_Galeria_opt.asp?acao=3&sF=<%= rsTemp("a00id") %>><img src="images/botao_apagar.gif" alt="Apagar este registro" width="16" height="16" border="0"></a></p>
                                            </td>
                                            <td align="center">
                                                <p><a href=pm_Galeria_fotos_opt.asp?acao=1&id=<%= rsTemp("a00id") %>><img src="images/botao_fotos.gif" alt="Inserir fotos para essa galeria" width="16" height="16" border="0"></a></p>
                                            </td>
                                            </td>
                                            <td align="center">
                                                <p><a href=pm_Galeria_fotos.asp?totext=<%= rsTemp("a00id") %>><img src="images/botao_vergaleria.gif" alt="Ver as fotos desta galeria" width="16" height="16" border="0"></a></p>
                                            </td>											
                                        </tr>
                                    </table>
                                </td>
                                <td align="Left">
                                    <p><font size="2" face="'Trebuchet MS', Verdana, Arial"><%= rsTemp("a03festa") %></font></p>
                                </td>
                                <td align="Left">
                                    <p><font size="2" face="'Trebuchet MS', Verdana, Arial"><%= rsTemp("a01local") %></font></p>
                                </td>
                                <td align="Center">
                                    <p><font size="2" face="'Trebuchet MS', Verdana, Arial"><%= rsTemp("a02data") %></font></p>
                                </td>
                                <td align="Center">
                                    <p><img src="galeria/<%= rsTemp(4) %>" alt="<%= rsTemp(3) %>" width="50" height="50" border="0"></font></p>
                                </td>
                            </tr>
<%
   rsTemp.MoveNext
   Loop
   rsTemp.Close
   Set rsTemp = Nothing
%>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td height="5" bgcolor="#0000ff">
        </td>
    </tr>
    <tr>
        <td class="tabelatitulo" height="10">
        </td>
    </tr>
</table>
		</td>
    </tr>
</table>
</body>
</html>
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
                        <p style="margin-left:5pt;"><span class="titulo">Informa��o do sistema...</span></p>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="128" align="center">
           <span class="titulo">Erro n� <%=err.Number%> <br><%=err.Source%><br><%= err.Description%></span>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="38" align="center">
           <span class="titulo"><a href="javascript:self.history.go(-1)">&lt;&lt; Voltar</a></span>
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
