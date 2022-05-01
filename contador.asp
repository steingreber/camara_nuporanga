<%
if Session("visitas") = "" Then
 set rsCounter = objConect.execute("SELECT * FROM tbcounter")
 visitas = rsCounter("acessos") + 1
 strsql = " UPDATE tbcounter SET "
 strsql = strsql & " acessos = " & visitas
 strsql = strsql & " where id = " & rsCounter("id")
 objConect.execute(strsql)
 Session("visitas") = rsCounter("acessos")
End If
 set iCounter = objConect.execute("Select * from tbcounter Where ID=1")
 Response.Write "Visitantes<br>" & iCounter(1)
 iCounter.Close
%>