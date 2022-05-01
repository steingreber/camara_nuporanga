<!--#include file="_cnx.asp"-->
<%
dim xAcao, xId
xAcao = Request("acao")
If xAcao = "ok" Then: xAcao = "0": Else: xAcao = "1"
xId   = Request("idm")
   sSel = sSel & "UPDATE CADASTROS SET "
   sSel = sSel & "A20LIBERADO = '" & xAcao & "' "
   sSel = sSel & "WHERE A00ID = " & xId
   Set objRS = objConect.Execute(sSel)
Response.Redirect "pm_CADASTROS.asp"
%>