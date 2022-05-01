<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<html>

<head>
<title><%= ititulo %></title>
<meta name="robots" content="None">
<script language="JavaScript" src="suport/suport.js"></script>
</head>

<body bgcolor="white" text="white" link="white" vlink="white" alink="white" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
		<TABLE WIDTH=778 BORDER=0 CELLPADDING=0 CELLSPACING=0>
                <TR>
<% If tipoimg = "0" Then %>
                    <TD height="133" width="778" background="config/imgtopo/<%= imgtopo %>">&nbsp;</TD>
<% Else %>
                    <TD height="133" width="778"><script>ConteudoFlash('config/imgTopo/<%= imgtopo %>','133','778');</script></TD>
<% End If %>
                </TR>
                <TR>
                    <td bgcolor=<%=corsel2%> width="778" align="right">
                        <p><font face="Trebuchet MS" size="2"><%=hoje%>&nbsp;</font></p>
                    </td>
                </TR>
		</TABLE>
</body>

</html>
