<!--#include file="_cnx.asp"-->
<%
dim xAcao, xId
xAcao = Request("acao")
xId   = Request("idm")
   sSel = sSel & "UPDATE menus SET "
   sSel = sSel & "a04exibir = '" & xAcao & "' "
   sSel = sSel & "WHERE a00id = " & xId
   Set objRS = objConect.Execute(sSel)
Response.Redirect "pm_menus.asp"
%>