<%
Dim xQual
xQual = Request.QueryString("xQual")
%>
<HTML>
  <HEAD>
    <TITLE>Prefeitura</TITLE>
  </HEAD>
    <FRAMESET ROWS=130,* FRAMEBORDER=1 FRAMESPACING=1 >
    <FRAME SRC="a_topoE.asp" SCROLLING=NO frameborder="0" noresize>
    <FRAME SRC="<%= xqual %>" SCROLLING=YES frameborder="0" noresize>
    </FRAMESET>
</HTML>