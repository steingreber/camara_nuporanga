<!--#include file="_cnx.asp"-->
<%
Dim sUser, sPass, vUser, vPass
sUser = Request("username")
sPass = Request("password")

	SQLPass = "Select * From login Where a02email = '" & sUser & "' AND a03senha = '" & sPass & "'"
	Set objRS = objConect.Execute(SQLPass)
	If objRS.EOF = True And objRS.BOF = True Then
		response.redirect "default.asp?mes=1"
	ElseIf Mid(Trim(objRS("a05permissoes")),23,1) = "0" Then
		response.redirect "default.asp?mes=2"
	End If

vUser = objRS(2)
vPass = objRS(3)
If sUser <> vUser or sPass <> vPass Then
   response.redirect "default.asp"
Else
   Session("xUser")   = True
   Session("xPass")   = True
   Session("usuario") = objRS(2)
   Session("nome")    = objRS(1)
   response.redirect "inicio.asp"
End If
objRS.Close
Set objRS = Nothing
%>

