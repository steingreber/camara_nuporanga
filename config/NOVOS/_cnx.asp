<%
'<------- /같같같같같같같같같같같같같\ ------->
'<------ / Gerando com PageMaster 2.0 \ ------>
'<----- /   http://www.gnove.com.br    \ ----->
'<----- \______________________________/ ----->

'NOME DO ARQUIVO...: _cnx.asp
'CRIADO EM.........: 19/6/2007 23:01:50
'-------------------------------------------------------------------------------
   Set objConect = Server.CreateObject("ADODB.Connection")
   strConn = "Provider=MSDataShape;DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & server.mappath("ROR543PRECAM4ROPKD\ROR944PRECAME993DE.mdb") & ";UID=;PWD="
   objConect.Open strConn
   session.LCID=1046

Dim asPathImage
'caminho da pasta das imagens
asPathImage = "" 'digite o caminho aqui - Server.MapPath("")
%>

