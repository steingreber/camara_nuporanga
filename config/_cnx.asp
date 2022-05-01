<%
'<------- /같같같같같같같같같같같같같\ ------->
'<------ / Gerando com PageMaster 2.5 \ ------>
'<----- /   http://www.gnove.com.br    \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: _cnx.asp
'CRIADO EM.........: 27/1/2007 14:02:48
'-------------------------------------------------------------------------------
   Set objConect = Server.CreateObject("ADODB.Connection")
   strConn = "Provider=MSDataShape;DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & server.mappath("ROR543PRECAM4ROPKD\ROR944PRECAME993DE.mdb") & ";UID=;PWD="
   objConect.Open strConn
   session.LCID=1046
%>

