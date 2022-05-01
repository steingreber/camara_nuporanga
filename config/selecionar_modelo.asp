<!--#include file="_cnx.asp"-->
<!--#include file="../config.asp"-->
<%
If Session("xUser") <> True And Session("xPass") <> True Then: response.redirect "vazou.asp"

	SQLPass = "Select * From login Where a02email = '" & Session("usuario") & "'"
	Set objRS = objConect.Execute(SQLPass)
	If Mid(objRS("a05permissoes"),39,1) = "0" Then
		response.redirect "entrada.asp?mes=1"
	End If
%>
<html>

<head>
<title>Selecione o modelo e cor para seu potal e clique gravar.</title>
<meta name="robots" content="none">
<link type="text/css" rel="stylesheet" href="pagemaster.css">
<SCRIPT language=Javascript>

function go2url() {
window.location = "entrada.asp";
}
</SCRIPT>
</head>

<body background="images/bg.gif" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#800080" alink="#FF0000">
<%
If Request.Form("cp_a24Template") = "" Then
%>
<table width="100%" height="100%" cellpadding="0" cellspacing="0">
    <tr>
        <td width="1057" align="center">
            <table width="500" height="53" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
                <tr>
                    <td width="500" colspan="3" align="center">
                        <p class="tabelatitulo"><font color="#000066">Selecione o modelo e cor para seu portal e clique gravar.</font></p>
                    </td>
                </tr>
                <tr>
                    <td width="500" colspan="3" align="left">
                        <input type="button" name="news_add" value=".::Cancelar::." class="botao" onClick="go2url()">
                    </td>
                </tr>
                <tr>
                    <td width="500" colspan="3" align="left">
						&nbsp;
                    </td>
                </tr>
                <tr>
                    <td width="261">
                        <p class="subtitulo">Modelo</p>
                    </td>
                    <td width="178">
                        <p class="subtitulo">Cor</p>
                    </td>
                    <td width="98">&nbsp;</td>
                </tr>
<%
pmCorSel = "#FFFFFF"
sLin = "Select * From TEMPLATES order by a01Template"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
Do While not objRs.EOF
If pmCorSel = "#E7E7E7" Then: pmCorSel = "#FFFFFF": Else pmCorSel = "#E7E7E7"
%>
			<form action="selecionar_modelo.asp" method="post">
                <tr>
                    <td width="261" height="30" bgcolor="<%= pmCorSel %>">
                            <p><input type="hidden" name="cp_a24Template" value="<%= objRs(0) %>"><span class="subtitulo"><%= objRs(1) %></span></p>
                    </td>
                    <td width="178" height="30" bgcolor="<%= pmCorSel %>">
           <select name="cp_a01corsite" size="1" class="ud_caixa">
<%
sLin2 = "Select * From CORES Where a06Tamplate="& objRS(0)
Set objRS2 = Server.CreateObject("ADODB.Recordset")
objRS2.Open sLin2, objConect, 3, 3
%>
                <option selected>Selecione</option>
<%Do While not objRs2.EOF%>
                <option value="<%= objRs2("a00id") %>"><%= objRs2("a08NomeCor") %></option>
<% objRs2.MoveNext
Loop
objRs2.close %>
                </select>
                    </td>
                    <td width="98" align="center" height="30" bgcolor="<%= pmCorSel %>">
                            <p><input type="submit" name="formbutton1" value="GRAVAR" class="ud_caixa"></p>
                    </td>
                </tr>
			</form>
<% objRs.MoveNext
Loop
objRs.close %>
            </table>
        </td>
    </tr>
</table>
<% Else
va_a24Template             = Trim(Request.Form("cp_a24Template"))
va_a01corsite              = Trim(Request.Form("cp_a01corsite"))

sNot = 1
   sSel = sSel & "UPDATE CONFIGURACOES SET "
   sSel = sSel & "a24Template = '" & va_a24Template & "', a01corsite = '" & va_a01corsite & "' "
   sSel = sSel & "WHERE a00idconfig = 1"
   response.write sSel
   Set objRS = objConect.Execute(sSel)
   response.redirect "entrada.asp"
End If
%>
<p>&nbsp;</p>
</body>

</html>