<%
sSql = "Select * From CONFIGURACOES"
   Set objRE = objConect.Execute(sSql)

iCores        = Trim(objRE(1))
iTitulo       = Trim(objRE(12)) & " Municipal de " & Trim(objRE(5)) & "-" & Trim(objRE(6))
iLegis        = Trim(objRE(16))
iCidade       = Trim(objRE(5)) & "/" &  Trim(objRE(6))
iCep          = Trim(objRE(7))
iFone  		  = Trim(objRE(8))
iMail    	  = Trim(objRE(10))
imgtopo 	  = Trim(objRE(14))
tipoimg 	  = Trim(objRE(15))
tvcamara	  = Trim(objRE(17))
clima         = Trim(objRE(19))
enquete       = Trim(objRE(20))
endereco_site = Trim(objRE(23))&"/"
iBannerG      = Trim(objRE(21))
iBannerP      = Trim(objRE(22))
iNoticias     = Trim(objRE(18))
iTemplate     = Trim(objRE(24))
iLargura      = Trim(objRE(26))
iAltura       = Trim(objRE(25))

Session("iEnd") = Trim(objRE(4)) & " - " & Trim(objRE(5)) & "/" & Trim(objRE(6)) & "<br> CEP " & Trim(objRE(7)) & "<br> FONE: " & Trim(objRE(8))

sSql2 = "Select * From CORES Where a00id="& iCores
   Set objRE2 = objConect.Execute(sSql2)

cCorTopoBase      = Trim(objRE2(1))
ca02corTopoTabela = Trim(objRE2(2))
ca03corCorpoTabela= Trim(objRE2(3))
Ca04corFonte      = Trim(objRE2(4))
ca05corLink       = Trim(objRE2(5))

function colore_busca( resultadoBusca, palavraBusca )   
  
    set objEreg = new regexp   
    with objEreg   
        .pattern    = "("& palavraBusca &")"  
        .ignorecase = true   
        .global     = true   
    end with   
  
    colore_busca = objEreg.replace( resultadoBusca , ""&_   
    "<font style=""color:RED; font-family:VERDANA; font-size:14px;""> $1 </font>" )   
       
end function       
'response.write colore_busca("Comentário: Aqui vem todas busca que o usuário efetuo inclusive o Comentário!.", "comentário")

'indica o ínicio do tempo a meia-noite 
if Cint(hour (time)) >= 0 Then 
mensagemtempo = "Boa-Madrugada" 
End if 

'indica o tempo até as 6 da manhã 
if Cint(hour (time)) >= 6 Then 
mensagemtempo = "Bom-Dia" 
End if 

'indica o tempo até as 12 da manhã 
if Cint(hour (time)) >= 12 Then 
mensagemtempo = "Boa-Tarde" 
End if 

'indica o tempo até as 18 da manhã 
if Cint(hour (time)) >= 18 Then 
mensagemtempo = "Boa-Noite" 
End if 

%>