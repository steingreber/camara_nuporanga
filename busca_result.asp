<%
'Dimension variables
Dim adoCon 'Variável da conexão da base de dados
Dim adoRec 'variável de Recordset da base de dados
Dim strAccessDB 'Holds the Access Database Name
Dim strSQL 'Database query sring
Dim intRecordPositionPageNum 'Holds the record position
Dim intRecordLoopCounter 'Loop counter for displaying the database records
Dim intTotalRecordsFound 'Holds the total number of records in the database
Dim intTotalNumPages 'holds the total number of pages in the database
Dim intLinkPageNum 'Holds the page number to be linked to
Dim strSearchKeywords 'Holds the keywords input by the user to be searched for
Dim strTitle 'Holds the URL Title
Dim strURL 'Holds the URL
Dim strDescription 'Holds the description of the URL
Dim sarySearchWord 'Holds the keywords for the URL
Dim intSQLLoopCounter 'Loop counter for the loop for the sql query
Dim intSearchWordLength 'Holds the length of the word to be searched
Dim blnSearchWordLenthOK 'Boolean set to false if the search word length is not OK
Dim intRecordDisplayFrom 'Holds the number of the search result that the page is displayed from
Dim intRecordDisplayTo 'Holds the number of the search result that the page is displayed to
Dim intRandomRecordNumber 'Holds a random number used to display a random URL

'Declare constants
' Mudar a seguinte linha ao número das entradas que você deseja ter em cada página e comprimento de palavra do miniumum 

Const intRecordsPerPage = 10 'Mudar este número à quantidade de entradas a ser indicadas em cada página 
Const intMinuiumSesrchWordLength = 2 'Mudar isto ao comprimento de palavra mínimo a ser procurarado sobre 

'Error handler
On error resume next

'Se isto for a primeira vez a página está indicada então a posição da página está ajustada para paginar 1
If Request.QueryString("PagePosition") = "" Then
intRecordPositionPageNum = 1

'A página foi indicada mais antes que o postion da página esteja ajustado assim ao número de posição Record 
Else
intRecordPositionPageNum = CInt(Request.QueryString("PagePosition"))
End If

'Ler dentro as palavras de busca a ser procuraradas 
sarySearchWord = Split(Trim(Request.QueryString("search")), " ")

'Ler dentro todas as palavras de busca em uma variável 
strSearchKeywords = Trim(Request.QueryString("search"))

'Substituir menos do que ou mais grande do que sinais com o equivalente do HTML (os batentes povoam Tag entrando do HTML)
strSearchKeywords = Replace(strSearchKeywords, "<", "&lt;")
strSearchKeywords = Replace(strSearchKeywords, ">", "&gt;")

'Initalise a variável do comprimento da busca da palavra 
blnSearchWordLenthOK = True

'Dar laços round para certificar-se de que cada palavra a ser procurarada tenha mais do que o comprimento de palavra mínimo a ser procurarado 
For intLoopCounter = 0 To UBound(sarySearchWord)

'Inicializar a variável do intSearchWordLength com o comprimento da palavra a ser procurarada 
intSearchWordLength = Len(sarySearchWord(intLoopCounter))

'Se o comprimento de palavra a ser procurarado fosse menos do que ou o igual ao comprimento de palavra mínimo ajustou então o blnWordLegthOK a falso
If intSearchWordLength <= intMinuiumSesrchWordLength Then 
blnSearchWordLenthOK = False 
End If
Next

'Initalise a variável do strSQL com uma indicação do SQL para perguntar a base de dados 
strSQL = "SELECT * FROM MATERIAL "

'Se o usuário selecionar para procurar algum intalise das palavras então a indicação do strSQL para procurarar por quaisquer palavras na base de dados 
If Request.QueryString("mode") = "anywords" Then

'Search for the first search word in the URL titles
strSQL = strSQL & "WHERE A01TITULO LIKE '%" & sarySearchWord(0) & "%'"

'Loop to search for each search word entered by the user
For intSQLLoopCounter = 0 To UBound(sarySearchWord)
strSQL = strSQL & " OR A01TITULO LIKE '%" & sarySearchWord(intSQLLoopCounter) & "%'"
strSQL = strSQL & " OR A06BUSCA LIKE '%" & sarySearchWord(intSQLLoopCounter) & "%'" 
strSQL = strSQL & " OR A02TEXTO LIKE '%" & sarySearchWord(intSQLLoopCounter) & "%'"
Next

End If

'If the user has selected to search for all words then intalise the strSQL statement to search for entries containing all the search words
If Request.QueryString("mode") = "allwords" Then

'Search for the first word in the URL titles
strSQL = strSQL & "WHERE A01TITULO LIKE '%" & sarySearchWord(0) & "%'"

'Loop to search the URL titles for each word to be searched
For intSQLLoopCounter = 1 To UBound(sarySearchWord)
strSQL = strSQL & " AND A01TITULO LIKE '%" & sarySearchWord(intSQLLoopCounter) & "%'"
Next

'OR if the search words are in the keywords 
strSQL = strSQL & " OR A06BUSCA LIKE '%" & sarySearchWord(0) & "%'"

'Loop to search the URL keywords for each word to be searched
For intSQLLoopCounter = 1 To UBound(sarySearchWord)
strSQL = strSQL & " AND A06BUSCA LIKE '%" & sarySearchWord(intSQLLoopCounter) & "%'"
Next

'Or if the search words are in the title 
strSQL = strSQL & " OR A02TEXTO LIKE '%" & sarySearchWord(0) & "%'"

'Loop to search the URL description for each word to be searched
For intSQLLoopCounter = 1 To UBound(sarySearchWord)
strSQL = strSQL & " AND A02TEXTO LIKE '%" & sarySearchWord(intSQLLoopCounter) & "%'"
Next

End If

'If the user has selected to see newly enetred URL's then order the search results by date decending
If Request.QueryString("mode") = "new" Then

'Order the search results by the date entered into the database decending
strSQL = strSQL & " ORDER By A01TITULO DESC"

'Else order the serch results by the number of click through hits decending 
Else
'Order the search results by the number of click through hits decending (most popular sites first)
strSQL = strSQL & " ORDER By A01TITULO DESC"
End If

'Response.Write strSQL

'Initialise the strAccessDB variable with the name of the Access Database
strAccessDB = "search_engine"

'Create a connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")
Set adoRec = Server.CreateObject("ADODB.Recordset")

'>----------------------------------------------------------
'Open connection to the database driver
'strCon = "DRIVER={Microsoft Access Driver (*.mdb)};"

'Open Connection to database
'strCon = strCon & "DBQ=" & server.mappath(strAccessDB)

'Query the database with the strSQL statement
'adoRec.Open strSQL, strCon, 3
%>
<!--#include file="_cnx.asp"-->
<%
'>----------------------------------------------------------

'Query the database with the strSQL statement
adoRec.Open strSQL, objConect, 3

'Ajustar o número dos registros à exposição em cada página pelo jogo constante no alto do certificado
adoRec.PageSize = intRecordsPerPage

'Get the page number record poistion to display from
adoRec.AbsolutePage = intRecordPositionPageNum

'Count the number of records found
intTotalRecordsFound = CInt(adoRec.RecordCount)

'Count the number of pages the search results will be displayed on calculated by the PageSize attribute set above
intTotalNumPages = CInt(adoRec.PageCount)

'Calculate the the record number displayed from and to on the page showing
intRecordDisplayFrom = (intRecordPositionPageNum - 1) * intRecordsPerPage + 1
intRecordDisplayedTo = (intRecordPositionPageNum - 1) * intRecordsPerPage + intRecordsPerPage
If intRecordDisplayedTo > intTotalRecordsFound Then intRecordDisplayedTo = intTotalRecordsFound

'Display the HTML table with the results status of the search or what type of search it is
Response.Write vbCrLf & " <br><table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"" align=""center"" bgcolor=""#dcdcdc"">"
Response.Write vbCrLf & " <tr>"

'Display that the URL is randomly generated 
If Request.QueryString("mode") = "random" Then
Response.Write vbCrLf & " <td class=""navtext""> Random URL.</td>"

'Display that we are showing a page of the latest URL's indexed
ElseIf Request.QueryString("mode") = "new" Then
Response.Write vbCrLf & " <td class=""navtext""> The " & intRecordsPerPage & " latest URL's Indexed.</td>"

'Display that one of the words entered was to short
ElseIf blnSearchWordLenthOK = False Then
Response.Write vbCrLf & " <td class=""navtext""> Procurado no site por <b><span class=""data"">" & strSearchKeywords & "</span></b>. &nbsp;&nbsp;&nbsp;Uma das palavras não conhecide.</td>"

'Display that there where no matching records found
ElseIf adoRec.EOF Then 
Response.Write vbCrLf & " <td class=""navtext""> Procurado no site por <b><span class=""data"">" & strSearchKeywords & "</span></b>. &nbsp;&nbsp;Desculpe, nada foi encontrado.</td>"

'Else Search went OK so display how many records found
Else 
Response.Write vbCrLf & " <td class=""navtext""> Procurado no site por <b><span class=""data"">" & strSearchKeywords & "</span></b>. &nbsp;&nbsp;&nbsp;Exibindo resultados " & intRecordDisplayFrom & " - " & intRecordDisplayedTo & " de " & intTotalRecordsFound & ".</td>" 
End If

'Close the HTML table with the search status
Response.Write vbCrLf & " </tr>"
Response.Write vbCrLf & " </table>"

'Display the various results

'HTML table to display the search results or an error if there are no results
Response.Write vbCrLf & " <br>" & vbCrLf 
Response.Write vbCrLf & " <table width=""99%"" border=""0"" cellspacing=""1"" cellpadding=""1"" align=""center"">"
Response.Write vbCrLf & " <tr>" 
Response.Write vbCrLf & " <td class=""data"">"

'Display a random URL
If Request.QueryString("mode") = "random" Then

'Randomise system timer
Randomize Timer

'Get a random number between 0 and highest number of records in database
intRandomRecordNumber = Int(RND * intTotalRecordsFound)

'Move to the choosen random record in the database
adoRec.Move intRandomRecordNumber

'Read in the values form the database
intSiteIDNo    = CInt(adoRec("A00ID")) 
strTitle       = "<font color=""#0000FF"">"&adoRec("A01TITULO")&"</font>"
strDescription = Left(adoRec("A02TEXTO"),160) & "..."
intSiteHits    = CInt(adoRec("A09EXIBIDA"))
strCateg       = adoRec("A05CATEG")
strSub         = adoRec("A11SUBCATEG")
strBlock       = adoRec("A08BLOQUEADO")
strURL         = endereco_site & "view_material.asp?id=" & intSiteIDNo & "&cat="& strCateg &"&sub="& strSub

'Display the randon URL details 
Response.Write vbCrLf & " <a href=""get_url.asp?SiteID=" & intSiteIDNo & """ target=""_blank""><span class=""navtext"">" & strTitle & "</span></a>"
Response.Write vbCrLf & " <br>"
Response.Write vbCrLf & " <span class=""navtext"">" & strDescription & "</span>"
Response.Write vbCrLf & " <br>"
Response.Write vbCrLf & " <font size=""1"" color=""#00CC99""><i>" & strURL & "&nbsp;&nbsp;-&nbsp;&nbsp;Hits&nbsp;" & intSiteHits & "</i></font>"
Response.Write vbCrLf & " <br><br>"

'Display error message if one of the words is to short
ElseIf blnSearchWordLenthOK = False And NOT Request.QueryString("mode") = "new" Then

'Write HTML displaying the error
Response.Write vbCrLf & " <span class=""navtext"">Sua busca - <b>" & strSearchKeywords & "</b> - Conteve uma palavra com " & intMinuiumSesrchWordLength & " letras ou menos."
Response.Write vbCrLf & " <br><br>"
Response.Write vbCrLf & " Sugestões:"
Response.Write vbCrLf & " <br>"
Response.Write vbCrLf & " <ul><li>Tente palavras maiores<li>Certificar-se que todas as palavras estão escritas corretamente.<li>Tente palavras diferentes.<li>Tente palavras mais usadas.</ul></span>"

'If no search results found then show an error message
ElseIf adoRec.EOF Then

'Write HTML displaying the error
Response.Write vbCrLf & " <span class=""navtext"">Sua busca - <b>" & strSearchKeywords & "</b> - não combinou com nenhuma das procuradas"
Response.Write vbCrLf & " <br><br>"
Response.Write vbCrLf & " Sugestões:"
Response.Write vbCrLf & " <br>"
Response.Write vbCrLf & " <ul><li>Certificar-se que todas as palavras estão escritas corretamente.<li>Tente palavras diferentes.<li>Tente palavras mais usadas.<li>Tentar poucas palavras.</ul></span>"

Else
'For....Next Loop to display the results from the database
For intRecordLoopCounter = 1 to intRecordsPerPage

'If there are no records left to display then exit loop
If adoRec.EOF Then Exit For

'Read in the values form the database
intSiteIDNo    = CInt(adoRec("A00ID")) 
strTitle       = "<font color=""#0000FF"">"&adoRec("A01TITULO")&"</font>"
strDescription = "<span class=""navtext"">" & Left(adoRec("A02TEXTO"),160) & "...</span>"
intSiteHits    = CInt(adoRec("A09EXIBIDA"))
strCateg       = adoRec("A05CATEG")
strSub         = adoRec("A11SUBCATEG")
strBlock       = adoRec("A08BLOQUEADO")
strURL         = endereco_site & "view_material.asp?id=" & intSiteIDNo & "&cat="& strCateg &"&sub="& strSub

If strBlock = "1" Then: strChave = "<img src=""imagens/bullet_chave.gif"" border=""0"" alt=""Para ver esta matéria precisa ser cadastrado no portal!""  align=""absmiddle"">"

'Display the details of the URLs found 
Response.Write vbCrLf & " <a href=""#"" onClick=""ajax('view_material.asp?id=" & intSiteIDNo & "&cat="& strCateg &"&sub="& strSub &"','conteudo')""><span class=""navtext"">" & strTitle & "</span></a>" & strChave
Response.Write vbCrLf & " <br>"
Response.Write vbCrLf & " " & colore_busca(strDescription,strSearchKeywords)
Response.Write vbCrLf & " <br>"
Response.Write vbCrLf & " <font size=""1"" color=""#00CC99""><i>" & strURL & "&nbsp;&nbsp;-&nbsp;&nbsp;Hits&nbsp;" & intSiteHits & "</i></font>"
Response.Write vbCrLf & " <br><br>"

'Move to the next record in the database
adoRec.MoveNext

'Loop back round 
Next
End If

'Close the HTML table displaying the results
Response.Write vbCrLf & " </td>"
Response.Write vbCrLf & " </tr>"
Response.Write vbCrLf & " </table>"

'Display an HTML table with links to the other search results if not showing latest 10, a random URL, or not under min word length
If NOT Request.QueryString("mode") = "new" And NOT Request.QueryString("mode") = "random" And blnSearchWordLenthOK = True Then

'Display an HTML table with links to the other search results
Response.Write vbCrLf & " <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">"
Response.Write vbCrLf & " <tr>"
Response.Write vbCrLf & " <td>"
Response.Write vbCrLf & " <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
Response.Write vbCrLf & " <tr>"
Response.Write vbCrLf & " <td width=""100%"" align=""center"">"

'If there are more pages to display then add a title to the other pages
If intRecordPositionPageNum > 1 or NOT adoRec.EOF Then
Response.Write vbCrLf & " Resultados:<br>"
End If

'If the page number is higher than page 1 then display a back link 
If intRecordPositionPageNum > 1 Then 
Response.Write vbCrLf & " <a href=""busca_ok.asp?PagePosition=" & intRecordPositionPageNum - 1 & "&search=" & Replace(strSearchKeywords, " ", "+") & "&mode=" & Request.QueryString("mode") &""" target=""_self"">&lt;&lt;&nbsp;Anterior</a>&nbsp; " 
End If

'If there are more pages to display then display links to all the search results pages
If intRecordPositionPageNum > 1 or NOT adoRec.EOF Then

'Loop to diplay a hyper-link to each page in the search results 
For intLinkPageNum = 1 to intTotalNumPages

'If the page to be linked to is the page displayed then don't make it a hyper-link
If intLinkPageNum = intRecordPositionPageNum Then
Response.Write vbCrLf & " " & intLinkPageNum
Else

Response.Write vbCrLf & " &nbsp;<a href=""busca_ok.asp?PagePosition=" & intLinkPageNum & "&search=" & Replace(strSearchKeywords, " ", "+") & "&mode=" & Request.QueryString("mode") & """ target=""_self"">" & intLinkPageNum & "</a>&nbsp; " 
End If
Next
End If

'If it is Not the End of the search results than display a next link 
If NOT adoRec.EOF then 
Response.Write vbCrLf & " &nbsp;<a href=""busca_ok.asp?PagePosition=" & intRecordPositionPageNum + 1 & "&search=" & Replace(strSearchKeywords, " ", "+") & "&mode=" & Request.QueryString("mode") & """ target=""_self"">Próximo&nbsp;&gt;&gt;</a>" 
End If

'Finsh HTML the table 
Response.Write vbCrLf & " </td>" 
Response.Write vbCrLf & " </tr>"
Response.Write vbCrLf & " </table>" 
Response.Write vbCrLf & " </td>"
Response.Write vbCrLf & " </tr>"
Response.Write vbCrLf & " </table>"
Response.Write vbCrLf & " <br>"

End If

'Close Server Objects 
Set adoCon = Nothing
Set adoCmd = Nothing
%>