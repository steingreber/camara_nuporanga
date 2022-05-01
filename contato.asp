
<%Response.ContentType = "text/html; charset=iso-8859-1"%>

<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->

<link type="text/css" rel="stylesheet" href="suport/suport.css">
<script language="JavaScript" src="suport/suport.js"></script>
<%If request.Form("cp1NOME") = "" Then%>
<form action="contato.asp" method="post" name="frmPageMaster" onSubmit="return Validator(this);">
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">FALE CONOSCO</font></td>
							</tr>
    <tr>
        <td width="451" colspan="2" class="navtext">
			Preencha os campos abaixo e envie sugestões, críticas ou comentários, pois sua opinião é fundamental e em breve entraremos em contato:
        </td>
    </tr>
        <tr>
            <td width="452" class="navtext" colspan="2">&nbsp;</td>
        </tr>
        <tr>
        <td width="150" class="navtext" height="27">Departamento</td>
            <td width="301" height="27">
           <select name="cp0DEPARTAMENTO" size="1" class="ud_caixa">
<%sLin = "Select * From CONTATO"
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open sLin, objConect, 3, 3
%>
                <option selected>Selecione</option>
<%Do While not objRs.EOF%>
                <option value="<%= objRs("A02EMAIL") %>"><%= objRs("A01NOME") %></option>
<% objRs.MoveNext
Loop
objRs.close %>
                </select>
            </td>
        </tr>
        <tr>
        <td width="150" class="navtext" height="27">Nome</td>
            <td width="301" height="27">
                <P><INPUT class=ud_caixa style="WIDTH:82%" name="cp1NOME"></P>
            </td>
        </tr>
        <tr>
        <td width="150" class="navtext" height="27">Cidade/Estado</td>
            <td width="301" height="27">
                <P><INPUT class=ud_caixa style="WIDTH:68%" name="cp2CIDADE">/<INPUT class=ud_caixa style="WIDTH:12%" name="cp3ESTADO"></P>
            </td>
        </tr>
        <tr>
        <td width="150" class="navtext" height="27">DDD-Fone</td>
            <td width="301" height="27">
                <P><INPUT class=ud_caixa style="WIDTH:12%" name="cp4DDD">-<INPUT class=ud_caixa style="WIDTH:68%" name="cp5FONE"></P>
            </td>
        </tr>
        <tr>
        <td width="150" class="navtext" height="27">E-Mail</td>
            <td width="301" height="27">
                <P><INPUT class=ud_caixa style="WIDTH:82%" name="cp6EMAIL"></P>
            </td>
        </tr>
        <tr>
        <td width="150" class="navtext">Mensagem</td>
            <td width="301">
                <P><TEXTAREA class=ud_caixa style="WIDTH:76%" name=pminfor rows="7" cols="46"></TEXTAREA></P>
            </td>
        </tr>
        <tr>
            <td width="451" colspan="2" height="46" align="center"><input type="image" name="btenviar" src="imagens/botao_enviar.gif" width="88" height="16" border="0"></td>
        </tr>
        <tr>
            <td width="451" height="46" align="center" colspan="2" class="navtext">
<%= Session("iEnd") %>
			</td>
        </tr>
</table>
</form>
<%  Else
    Dim sMsgErr
    Sub ProcessaPagina()
    Dim va1NOME
    Dim va2CIDADE
    Dim va3ESTADO
    Dim va4DDD
    Dim va5FONE
    Dim va6EMAIL

	va00DEPARTAMENTO		  = Request.Form("cp0DEPARTAMENTO")
    va1NOME                   = Request.Form("cp1NOME")
    va2CIDADE                 = Request.Form("cp2CIDADE")
    va3ESTADO                 = Request.Form("cp3ESTADO")
    va4DDD                    = Request.Form("cp4DDD")
    va5FONE                   = Request.Form("cp5FONE")
    va6EMAIL                  = Request.Form("cp6EMAIL")
    vaMENSAGEM                = Replace(Request.Form("pminfor"),chr(13),"<BR>",1)

    snavtext = snavtext & "<font face='Courier New' color='#333399'>" & vbCrLf
    snavtext = snavtext & "<b><div align=center>Mensagem recebida do site</b></div><br><br>" & vbCrLf
    snavtext = snavtext & "NOME...........: " & va1NOME & "<br>" & vbCrLf
    snavtext = snavtext & "CIDADE/ESTADO..: " & va2CIDADE & "/" & va3ESTADO & "<br>" & vbCrLf
    snavtext = snavtext & "DDD-FONE.......: " & va4DDD & "-" & va5FONE & "<br>" & vbCrLf
    snavtext = snavtext & "EMAIL..........: " & va6EMAIL & "<br>" & vbCrLf
    snavtext = snavtext & "<br><b>MENSAGEM</b> <br>" & vbCrLf
    snavtext = snavtext & vaMensagem & "<br>" & vbCrLf
    snavtext = snavtext & "<hr><br>" & vbCrLf
    snavtext = snavtext & "<font face=Verdana size=1 color='#0000FF'><a href='http://www.gnove.com.br'>Gentileza www.gnove.com.br</a>" & vbCrLf
    snavtext = snavtext & ""

'    On Error Resume Next
    Set pmMail = Server.CreateObject("CDONTS.NewMail")
    pmMail.mailFormat = 0
    pmMail.bodyFormat = 0
    pmMail.From = va6EMAIL
    pmMail.To   = va00DEPARTAMENTO
    pmMail.Subject = "CONTATO DO PORTAL "& iTitulo
    pmMail.Body = snavtext
    pmMail.Send
    Set pmMail = Nothing
If Err <> 0 Then
    sMsgErr = "Ocorreu o seguinte erro ao tentar enviar o e-mail:" & Err.Description
End If
    On Error GoTo 0
    End Sub
    ProcessaPagina
	Response.Redirect "mensagens.asp?id=1"
	End If
%>
<!-- fim -->