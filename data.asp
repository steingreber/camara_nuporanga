<%
Dim dia
Dim mes
dim ano
Dim semana
Dim etenso
Dim DiaSemana
Dim NomeMes
Dim hoje

dia=day(Date)
mes=month(Date)
ano=year(Date)
semana=weekday(Date)

Select Case Semana
Case 1
DiaSemana = "Domingo"
Case 2
DiaSemana = "Segunda"
Case 3
DiaSemana = "Terça"
Case 4
DiaSemana = "Quarta"
Case 5
DiaSemana = "Quinta"
Case 6
DiaSemana = "Sexta"
Case 7
DiaSemana = "Sabado"
End Select

Select Case mes
Case 1
NomeMes = "Janeiro"
Case 2
NomeMes = "Fevereiro"
Case 3
NomeMes = "Março"
Case 4
NomeMes = "Abril"
Case 5
NomeMes = "Maio"
Case 6
NomeMes = "Junho"
Case 7
NomeMes = "Julho"
Case 8
NomeMes = "Agosto"
Case 9
NomeMes = "Setembro"
Case 10
NomeMes = "Outubro"
Case 11
NomeMes = "Novembro"
Case 12
NomeMes = "Dezembro"
End Select

hoje = DiaSemana & ", " & dia & " de " & NomeMes & " de " & ano & "&nbsp;"
%>