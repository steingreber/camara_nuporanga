<!--#include file="_cnx.asp"-->
<%
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"

	SQLPass = "Select * From login Where a02email = '" & Session("usuario") & "'"
	Set objRS = objConect.Execute(SQLPass)
	If Mid(objRS("a05permissoes"),2,1) = "0" Then
		response.redirect "entrada.asp?mes=1"
	End If
%>
<html>
<head>
<title>NOTICIAS</title>
<meta name="robots" content="none">
<link rel="stylesheet" href="pagemaster.css">
<link type="text/css" rel="stylesheet" href="../suport/suport.css">
<script language="JavaScript" src="../suport/suport.js"></script>
<script language=JavaScript>
<!--
function confirm_delete() {
  return confirm ("Voc? realmente deseja excluir este registro?")
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
On error resume next
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
   SQL = "Select [a00id], [a01data], [a02titulo] From [noticias] Order By [a00id] Desc"
Else
   If Trim(sTipo) = "Like" Then
   SQL = "Select * From [noticias] Where " & sCampo & " " & sTipo & "'%" & sValor & "%'"
   Else
   SQL = "Select * From [noticias] Where " & sCampo & " " & sTipo & "'" & sValor & "'"
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
        <td height="5" colspan="3" bgcolor="#0000ff">
        </td>
    </tr>
<tr>
	<td colspan="3" class="tabelatitulo">
		&nbsp;<span class="aFonte">&raquo;&raquo;&nbsp;NOT?CIAS&nbsp;&laquo;&laquo;</span>
	</td>
</tr>
    <tr>
          <form name="frmBusca" method="post" action=pm_noticias.asp>
        <td height="34" align="center" colspan="3" class="tabelatitulo">
               <p><select name="q" size=1 class="ud_caixa">
               <option selected value="[a01data]">DATA</option>
               <option value="[a02titulo]">TITULO</option>
               </select> 
               <select name="t" class="ud_caixa">
               <option selected value=" Like ">que contenha</option>
               <option value=" = ">igual a...</option>
               </select>
               <input type="text" name="totext" size=15 class="ud_caixa">
                <input type="submit" name="Search" value="Buscar" class="botao"></p>
        </td>
       </form>
    </tr>
    <tr>
        <td width=494 height=20 align="center" class="tabelatitulo">
            <p><a href="pm_noticias_opt.asp?acao=1"><span class="aFonte">INSERIR NOVA NOT?CIA</span></a></p>
        </td>
        <td width="90" align="right" height=20 class="tabelatitulo">
               <span class="aFonte">NAVEGADOR:</span></td>
        <td width="269" align="center" height="20" class="tabelatitulo">
<span class="aFonte"><B>
<%
  Response.Write "<A HREF=""pm_noticias.asp?Page=1""><img src='images/p2.gif' alt=Primeiro width=16 height=16 border=0> </A>"
if Current_Page > 1 then
  Response.Write "<A HREF=""pm_noticias.asp?Page=" & (Current_Page - 1) & """><img src='images/p1.gif' alt=Anterior width=16 height=16 border=0> </A> " & vbCrLfelse
  Response.Write ""
end if
if Current_Page < Page_Count then
  Response.Write "<A HREF=""pm_noticias.asp?Page=" & (Current_Page + 1) & """><img src='images/p3.gif' alt=Pr?ximo width=16 height=16 border=0> </A>"
else
  Response.Write ""
end if
  Response.Write "<A HREF=""pm_noticias.asp?Page=" & Page_Count & """><img src='images/p4.gif' alt=?ltimo width=16 height=16 border=0>   </A>"
%>
</B></FONT></td>
    </tr>
    <tr>
        <td height=5 colspan=3 class="tabelaitem">
        </td>
    </tr>
    <tr>
        <td colspan=3 align="center" valign="top">
            <table width="100%" border="1" cellspacing="0" align="center" bordercolorlight="#808080" bordercolordark="#FFFFFF">
                <tr>
                    <td>
                        <table align="center" cellpadding="0" cellspacing="1" width="100%">
                            <tr>
                                <td width="63" align="center" class="tabelatitulo">
                                    <p><span class="aFonte">OP??ES</span></p>
                                </td>
                                <td align="center" class="tabelatitulo">
                                    <p><span class="aFonte">DATA</span></p>
                                </td>
                                <td align="center" class="tabelatitulo">
                                    <p><span class="aFonte">TITULO</span></p>
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
                                <td width="63">
                                    <table cellpadding="0" cellspacing="0" width="100%" align="center">
                                        <tr>
                                            <td align="center">
                                                <p><a href="pm_noticias_opt.asp?acao=2&sF=<%= rsTemp("a00id") %>"><img src="images/botao_editar.gif" alt="EDITAR REGISTRO" width="16" height="16" border="0"></a></p>
                                            </td>
                                            <td align="center">
                                                <p><a onclick="return confirm_delete()" href=pm_noticias_opt.asp?acao=3&sF=<%= rsTemp("a00id") %>><img src="images/botao_apagar.gif" alt="EXCLUIR REGISTRO" width="16" height="16" border="0"></a></p>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                                <td align="Center">
                                    <p><font size="2" face="Trebuchet MS"><%= rsTemp("a01data") %></font></p>
                                </td>
                                <td align="Left">
                                    <p><font size="2" face="Trebuchet MS"><%= rsTemp("a02titulo") %></font></p>
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
        <td height="5" colspan="3" bgcolor="#0000ff">
        </td>
    </tr>
    <tr>
        <td colspan="3" class="tabelatitulo" height="10">
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
                        <p style="margin-left:5pt;"><span class="titulo">Informa??o do sistema...</span></p>
                    </td>
                </tr>
                <tr>
                    <td width="571" height="128" align="center">
           <span class="titulo">Erro n? <%=err.Number%> <br><%=err.Source%><br><%= err.Description%></span>
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

