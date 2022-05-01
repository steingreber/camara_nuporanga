<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<%
va_A19DATACAD        = date
va_A01NOME           = Trim(Request("fu01Nome"))
va_A02ENDERECO       = Trim(Request("fu14EndRes"))
va_A03NUMERO         = Trim(Request("fu15Numero"))
va_A04BAIRRO         = Trim(Request("fu17Bairro"))
va_A05COMPLEMENTO    = Trim(Request("fu16Complemento"))
va_A06CIDADE         = Trim(Request("fu19Cidade"))
va_A07ESTADO         = Trim(Request("fu20Estado"))
va_A23CEP            = Trim(Request("fu18CEP"))
va_A08FONE           = Trim(Request("fu04FoneContato"))
va_A09CELULAR        = Trim(Request("fu21FoneRes"))
va_A10EMAIL          = Trim(Request("fu30username"))
va_A17SENHA          = Trim(Request("fu31senha"))
va_A11CPF            = Trim(Request("fu06CPF"))
va_A12RG             = Trim(Request("fu07Identidade"))
va_A13ESTADOCIVIL    = Trim(Request("fu05EstadoCivil"))
va_A14SEXO           = Trim(Request("fu06Sexo"))
va_A15PROFISSAO      = Trim(Request("fu07Profissao"))
va_A18DATANASC       = Trim(Request("fu02Data_Nascto"))
va_A20LIBERADO       = "0"
va_A22PLANO          = "0"

va_A21LIBERADOATEH   = DateAdd("d",iPrazo+5,date)

   sSel = sSel & "INSERT INTO CADASTROS(A19DATACAD, A01NOME, A02ENDERECO, A03NUMERO, A04BAIRRO, A05COMPLEMENTO, A06CIDADE, A07ESTADO, A23CEP, A08FONE, A09CELULAR, A10EMAIL, A17SENHA, A11CPF, A12RG, A13ESTADOCIVIL, A14SEXO, A15PROFISSAO, A18DATANASC, A20LIBERADO, A21LIBERADOATEH, A22PLANO)"
   sSel = sSel & "VALUES ('" & va_A19DATACAD & "', '" & va_A01NOME & "', '" & va_A02ENDERECO & "', '" & va_A03NUMERO & "', '" & va_A04BAIRRO & "', '" & va_A05COMPLEMENTO & "', '" & va_A06CIDADE & "', '" & va_A07ESTADO & "', '" & va_A23CEP & "', '" & va_A08FONE & "', '" & va_A09CELULAR & "', '" & va_A10EMAIL & "', '" & va_A17SENHA & "', '" & va_A11CPF & "', '" & va_A12RG & "', '" & va_A13ESTADOCIVIL & "', '" & va_A14SEXO & "', '" & va_A15PROFISSAO & "', '" & va_A18DATANASC & "', '" & va_A20LIBERADO & "', '" & va_A21LIBERADOATEH & "', '" & va_A22PLANO & "')"
   Set objRS = objConect.Execute(sSel)
   'objRS.Close
   'Set objRS = Nothing

    sms_href = sms_href & "<font face='Courier New' color='#333399'>" & vbCrLf
    sms_href = sms_href & "<b><div align=center>Mensagem de confirmação de cadastro!</b></div><br><br>" & vbCrLf
    sms_href = sms_href & "NOME...........: " & va_A01NOME & "<br>" & vbCrLf
    sms_href = sms_href & "CIDADE/ESTADO..: " & va_A06CIDADE & "/" & va_A07ESTADO & "<br>" & vbCrLf
    sms_href = sms_href & "DDD-FONE.......: " & va_A08FONE & "<br>" & vbCrLf
    sms_href = sms_href & "EMAIL..........: " & va_A10EMAIL & "<br>" & vbCrLf
	sms_href = sms_href & "SENHA..........: " & va_A17SENHA & "<br>" & vbCrLf
	'sms_href = sms_href & "PLANO DE ACESSO: " & Iplano & "<br>" & vbCrLf
    sms_href = sms_href & "<br><b>MENSAGEM</b> <br>" & vbCrLf
    sms_href = sms_href & vaMensagem & "<br>" & vbCrLf
    sms_href = sms_href & "<hr><br>" & vbCrLf
    sms_href = sms_href & "<font face=Verdana size=1 color='#0000FF'><a href='http://www.gnove.com.br'>Gentileza www.gnove.com.br</a>" & vbCrLf
    sms_href = sms_href & ""

    Set pmMail = Server.CreateObject("CDONTS.NewMail")
    pmMail.mailFormat = 0
    pmMail.bodyFormat = 0
    pmMail.From = iMail
    pmMail.To   = va_A10EMAIL
    pmMail.Subject = "Contato do portal " & iTitulo
    pmMail.Body = sms_href
    pmMail.Send
    Set pmMail = Nothing

	Response.Redirect "mensagens.asp?mensagem=1"
%>


