<!--#include file="_cnx.asp"-->
<%
Dim sUser, sPass, vUser, vPass
sUser = Request("email")
sPass = Request("senha")

SQL = "Select A10EMAIL, A17SENHA, A21LIBERADOATEH, A14SEXO, A01NOME, A20LIBERADO From CADASTROS Where A10EMAIL='"& sUser &"' AND A17SENHA='"& sPass &"'"
   Set objRS = objConect.Execute(SQL)

If objRS.BOF = True And objRS.EOF = True Then
	Response.Redirect "mensagens.asp?id=4"
ElseIf date > objRS("A21LIBERADOATEH") Then
	Response.Redirect "mensagens.asp?id=3"
ElseIf objRS("A20LIBERADO")=0 Then
	Response.Redirect "mensagens.asp?id=5"
Else
   Session("vUser") = UCase(objRS("A01NOME"))
   Session("iUser") = True
   Session("iPass") = True
   Session("xSexo") = objRS("A14SEXO")
   response.redirect "default.asp"
End If
%>

