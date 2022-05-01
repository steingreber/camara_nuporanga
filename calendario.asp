<link type="text/css" rel="stylesheet" href="suport/suport.css">
<script src="suport/flash.js"></script>

<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
<%
Dim auxiliarF(31)
    sqlF = "SELECT * FROM CALENDARIO Where Year(o02data)='" & Year(date) & "'"
    Set rsTempF = Server.CreateObject("ADODB.Recordset")
    rsTempF.Open sqlF, objConect,3,3	

'mesobrF = month(rstempF("o02data"))
'response.write ("MES: ")
'response.write mesobrF

if rsTempF.bof = false and rsTempF.eof = false then
  Do while not rsTempF.eof
	mesobrF = month(rstempF("o02data"))
	anoobrF = year(rstempF("o02data"))
'	response.write mesobrF & "/" & anoobrF & "<BR>"
    if mesobrF = month(date) then
	dtObrigacao = rstempF("o02data")  
	diaobrF     = day(dtObrigacao)

Select case diaobrF
	Case 1  auxiliarF(1)  = "Sim"
	Case 2  auxiliarF(2)  = "Sim"
	Case 3  auxiliarF(3)  = "Sim"
	Case 4  auxiliarF(4)  = "Sim"
	Case 5  auxiliarF(5)  = "Sim"
	Case 6  auxiliarF(6)  = "Sim"
	Case 7  auxiliarF(7)  = "Sim"
	Case 8  auxiliarF(8)  = "Sim"
	Case 9  auxiliarF(9)  = "Sim"
	Case 10 auxiliarF(10) = "Sim"
	Case 11 auxiliarF(11) = "Sim"
	Case 12 auxiliarF(12) = "Sim"
	Case 13 auxiliarF(13) = "Sim"
	Case 14 auxiliarF(14) = "Sim"
	Case 15 auxiliarF(15) = "Sim"
	Case 16 auxiliarF(16) = "Sim"
	Case 17 auxiliarF(17) = "Sim"
	Case 18 auxiliarF(18) = "Sim"
	Case 19 auxiliarF(19) = "Sim"
	Case 20 auxiliarF(20) = "Sim"
	Case 21 auxiliarF(21) = "Sim"
	Case 22 auxiliarF(22) = "Sim"
	Case 23 auxiliarF(23) = "Sim"
	Case 24 auxiliarF(24) = "Sim"
	Case 25 auxiliarF(25) = "Sim"
	Case 26 auxiliarF(26) = "Sim"
	Case 27 auxiliarF(27) = "Sim"
	Case 28 auxiliarF(28) = "Sim"
	Case 29 auxiliarF(29) = "Sim"
	Case 30 auxiliarF(30) = "Sim"
	Case 31 auxiliarF(31) = "Sim"
End Select
	End if
    rstempF.movenext
	Loop
End If


'Lógica para construir o calendário
Session.LCID = 1046 'BRASIL
Dim URL,Dia, Mes, Ano, Agora, PrimeiroDiaMes, UltimoDiaMes 
Dim Inicio,Fim, Start, TheEnd,i, j

URL= Request.ServerVAriables("SCRIPT_NAME")
 
if (Request.QueryString("Mes") <> Month(now))AND(Request.QueryString("Mes") <> "") Then
	Mes = CInt(Request.QueryString("Mes"))'força que o resultado seja um inteiro
else
	Mes = Month(now)
end if

Dia   =Day(now)
Ano   = Year(now)
Agora = DateSerial(Ano, mes, dia)

PrimeiroDiaMes = DateSerial(Year(Now),Mes,1)
UltimoDiaMes   = DateSerial(Year(Now),Mes+1,1-1)
Inicio         = ABS(1 - WeekDay(PrimeiroDiaMes))
Fim            = 7 - WeekDay(UltimoDiaMes)
Start          = 1-Inicio
TheEnd         = Day(UltimoDiaMes) + Fim 
J              = 1
%>

                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">CALENDÁRIO DE OBRIGAÇÕES</font></td>
							</tr>
							<tr>
							<td class="navtext">Clique no dia desejado para ver as obrigações!</td>
							</tr>
						</table>

<br><br>
<TABLE BORDER=0 WIDTH=280 CELLPADDING=0 ALIGN=center>
	<TR>
	<TD WIDTH="1%" VALIGN=top>
<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0 WIDTH="100%">

	<TR>
<TD NOWRAP COLSPAN=4 ALIGN=CENTER BGCOLOR=dcdcdc>
	<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD>
<A HREF="<%=URL&"?mes="&(mes-1)%>"></A>
		</TD>
		<td valign="middle" nowrap class="tittext">
			<strong><%Response.Write Ucase(MonthName(Month(Agora))) & " de " & Year(date)%></strong>
		</td>
		<TD>
<A HREF="<%=URL&"?mes="&(mes+1)%>"></A>
		</TD>
	</TR>
	</TABLE>
</TD>
</TR>
<tr>
	<TD COLSPAN=4 align="Center">
		<TABLE  CELLSPACING=0 CELLPADDING=0 BORDER=0>
		<TR>
			<TD COLSPAN=1 align="center" width='44' height='31' bgcolor="<%= ca02corTopoTabela %>" class="tittext"><b>DOM</b></TD>
			<TD COLSPAN=1 align="center" width='44' height='31' bgcolor="<%= ca02corTopoTabela %>" class="tittext"><b>SEG</b></TD>
			<TD COLSPAN=1 align="center" width='44' height='31' bgcolor="<%= ca02corTopoTabela %>" class="tittext"><b>TER</b></TD>
			<TD COLSPAN=1 align="center" width='44' height='31' bgcolor="<%= ca02corTopoTabela %>" class="tittext"><b>QUA</b></TD>
			<TD COLSPAN=1 align="center" width='44' height='31' bgcolor="<%= ca02corTopoTabela %>" class="tittext"><b>QUI</b></TD>
			<TD COLSPAN=1 align="center" width='44' height='31' bgcolor="<%= ca02corTopoTabela %>" class="tittext"><b>SEX</b></TD>
			<TD COLSPAN=1 align="center" width='44' height='31' bgcolor="<%= ca02corTopoTabela %>" class="tittext"><b>SAB</b></TD>
		</TR>
		</Table>
          </TD>
</TR>
<TR>
	<TD COLSPAN=4 align="Center">
	<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=1>
		
	<TR bgcolor="<%= ca03corCorpoTabela %>">

<%
For i= Start to TheEnd  
auxF = Day(DateSerial(Year(Now),Mes,i))
  if (auxiliarF(auxF) = "Sim") and (mesobrF = Month(DateSerial(Year(Now),Mes,i))) and (anoobrF = year(DateSerial(Year(Now),Mes,i))) Then
	Response.Write "<TD align= center width='44' height='44' bgcolor='"& ca03corCorpoTabela &"'><a href=javascript:ajax('calendario_view.asp?dia=" & Day(DateSerial(Year(Now),Mes,i)) & "&mes=" & Month(DateSerial(Year(Now),Mes,i)) & "&ano=" & year(DateSerial(Year(Now),Mes,i)) & "','calendario')><B><font color='blue'><strong>"& Day(DateSerial(Year(Now),Mes,i)) & "</strong></font></B></a></TD>"
  Elseif WeekDay(DateSerial(Year(Now),Mes,i)) = 1 Then
	Response.Write "<TD align= center width='44' height='44'><FONT face=verdana size=3 COLOR=red>"& Day(DateSerial(Year(Now),Mes,i)) & "</font></TD>"
  Elseif (i = day(now)) And (Mes=Month(now)) Then
	Response.Write "<TD align= center width='44' height='44' bgcolor='#ff0000'><FONT face=verdana size=3 COLOR=white><b>"& Day(DateSerial(Year(Now),Mes,i)) & "</b></font></TD>"
  Elseif (i<1) or (i > Day(UltimoDiaMes)) Then
	Response.Write "<TD align= center width='44' height='44'><FONT face=verdana size=3 COLOR=white><B>"& Day(DateSerial(Year(Now),Mes,i)) & "</B></TD>"
  Else
	Response.Write "<TD align= center width='44' height='44'><FONT face=verdana size=3 COLOR=""#000000"">"& Day(DateSerial(Year(Now),Mes,i)) & "</TD>"
  End if

  If j = 7 then
     J = 0
	 Response.Write "</tr>"
	 Response.Write "<TR bgcolor='"& ca03corCorpoTabela &"'>"
End if
  J = j + 1
Next 
%>
      </TR>
	  <tr>
	  <td colspan="7" align="center" bgcolor="#FFFFFF" class="tittext">
	  <font color="#FF0000">Dia Atual</font> - <font color="#33ccff">Obrigações do dia</font>
	  </td>
	  </tr>
      </Table>
      </TD>
</tr></Table></Table>
<div id="calendario"></div>