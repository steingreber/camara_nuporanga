<%
Dim sUser, sPass, vUser, vPass
sUser = Request("username")
sPass = Request("password")
vUser = "DFG"
vPass = "FG"
If sUser <> vUser or sPass <> vPass Then
   response.redirect "default.asp"
Else
   Session("xUser") = True
   Session("xPass") = True
   response.redirect "pm_CATEGORIAS.asp"
End If
%>

