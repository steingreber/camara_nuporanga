<%
	'==============================================================
	' RSS/RDF Syndicate Reader v0.95
	' http://www.kattanweb.com/webdev
	'--------------------------------------------------------------
	' Copyright(c) 2002, KattanWeb.com
	'
	' Change Log:
	'--------------------------------------------------------------
	'==============================================================

Const rssInit  = 1
Const rssError = 2
Const rssBadRSS= 3
Const rssOK    = 0

class kwRSS_reader
	Private Items()
	Private CurrentItem, TotalItems
	Public ChannelRSSURI, ChannelURL, ChannelTitle, ChannelDesc, ChannelLanguage
	Public ImageTitle, ImageLink, ImageURL
	Public TextInputURL, TextInputTitle, TextInputDesc, TextInputName
	Private Status

    '>>>>>>>> Setup Initialize event, called automtially when creating an instant of this class using
	'	      Set rss = new kwRSS_reader
	Private Sub Class_Initialize
		CurrentItem = -1
		TotalItems  = -1
		Redim Items(2, 10)		'1st dimension = item's title/link/desc, 2nd dimension the item number
		Status = rssInit
	End Sub

    '>>>>>>>> Setup Terminate event, called automtially when killing an instant of this class using
	'	      Set rss = nothing
	Private Sub Class_Terminate
		Erase Items
	End Sub
 
    '>>>>>>>> Load an RSS/RDF file and process it.
	Public Function ParseLocation(URL)
		ChannelRSSURI = URL
		set xmlObj = Server.CreateObject("Msxml2.DOMDocument.3.0")
		set xmlhttp= Server.CreateObject("Msxml2.XMLHTTP.3.0")

			xmlObj.validateOnParse = false
			xmlObj.async = false
			xmlObj.preserveWhiteSpace = false

			xmlhttp.open "GET", ChannelRSSURI, False
			xmlhttp.send
			rssXML_Data = xmlhttp.responseXML.xml
			' --------- PHP-Nuke & PostNuke compatabilaty --------------------------------
			rssXML_Data = replace(rssXML_Data, "<!DOCTYPE", "<!--DOCTYPE")
			rssXML_Data = replace(rssXML_Data, "rss-0.91.dtd"">", "rss-0.91.dtd""-->")
			' ----------------------------------------------------------------------------
			xmlObj.loadXML(rssXML_Data)
			If xmlObj.parseError.errorCode = 0 then
				ValidLocation = true
			else
				ValidLocation = false
			end if
		set xmlhttp = nothing

		if not ValidLocation then
			Status = rssBadRSS
			Exit Function
		end if

		set rootNode = xmlObj.selectSingleNode("rdf:RDF")
		if rootNode is nothing then
			set rootNode = xmlObj.selectSingleNode("rss")
			if rootNode is nothing then
				Status = rssError
			else
				Reader rootNode, 0.91
			end if
		else
			Reader rootNode, 1.0
		end if		

		set rootNode = nothing
		set xmlObj = nothing
		Status = rssOK
	End Function

    '>>>>>>>> Private sub to read the RSS/RDF according to its version
	Private Sub Reader(rootNode, ver)
		itemNum = -1

		set SingleNode = rootNode.selectSingleNode("//channel/title")
		if Not SingleNode is nothing then ChannelTitle = SingleNode.text

		set SingleNode = rootNode.selectSingleNode("//channel/link")
		if Not SingleNode is nothing then ChannelURL = SingleNode.text

		set SingleNode = rootNode.selectSingleNode("//channel/description")
		if Not SingleNode is nothing then ChannelDesc = SingleNode.text

		set SingleNode = rootNode.selectSingleNode("//channel/language")
		if Not SingleNode is nothing then ChannelLanguage = SingleNode.text

		if ver = 1 then
			set child = rootNode.selectSingleNode("image")
		else
			set child = rootNode.selectSingleNode("//channel/image")
		end if
			if not child is nothing then
				set SingleNode = child.selectSingleNode("title")
				if Not SingleNode is nothing then ImageTitle = SingleNode.text

				set SingleNode = child.selectSingleNode("link")
				if Not SingleNode is nothing then ImageLink = SingleNode.text

				set SingleNode = child.selectSingleNode("url")
				if Not SingleNode is nothing then ImageURL = SingleNode.text
			end if
		set child = nothing

		if ver = 1 then
			set child = rootNode.selectSingleNode("textinput")
		else
			set child = rootNode.selectSingleNode("//channel/textinput")
		end if
			if not child is nothing then
				set SingleNode = child.selectSingleNode("title")
				if Not SingleNode is nothing then TextInputTitle = SingleNode.text

				set SingleNode = child.selectSingleNode("description")
				if Not SingleNode is nothing then TextInputDesc = SingleNode.text

				set SingleNode = child.selectSingleNode("name")
				if Not SingleNode is nothing then TextInputName = SingleNode.text

				set SingleNode = child.selectSingleNode("link")
				if Not SingleNode is nothing then TextInputURL = SingleNode.text
			end if
		set child = nothing

		set children = rootNode.selectNodes("//item")
		TotalItems = children.length
			for each child in children
				itemNum = itemNum + 1
				if itemNum > ubound(Items, 2) then
					Redim Preserve Items(2, ubound(Items, 2) + 5)
				end if
				for each ItemChild in child.ChildNodes
					select case ItemChild.baseName
						case "title"		Items(0, itemNum) = ItemChild.text
						case "link"			Items(1, itemNum) = ItemChild.text
						case "description"	Items(2, itemNum) = ItemChild.text
					end select
				next
			next
		if TotalItems > 0 then CurrentItem = 0
	End Sub

    '>>>>>>>> Returns the title of the the current item
	Public Function GetTitle()
		GetTitle = Items(0, CurrentItem)
	End Function

    '>>>>>>>> Returns the url/link of the the current item
	Public Function GetLink()
		GetLink = Items(1, CurrentItem)
	End Function

	'>>>>>>>> Returns the description of the the current item
	Public Function GetDesc()
		GetDesc = Items(2, CurrentItem)
	End Function
	
    '>>>>>>>> Goes to the next item
	Public Function MoveNext
		CurrentItem = CurrentItem + 1
	End Function

    '>>>>>>>> Goes to the first item
	Public Function FirstItem
		if TotalItems > 0 then
			CurrentItem = 0
		else
			CurrentItem = -1
		end if
	End Function

    '>>>>>>>> Checks if the current location is a valid item or not
	Public Function ValidItem
		if CurrentItem > -1 and CurrentItem < TotalItems then
			ValidItem = true
		else
			ValidItem = false
		end if			
	End Function

    '>>>>>>>> Checks if we are at EOF or not
	Public Function EOF
		if CurrentItem < TotalItems then
			EOF = false
		else
			EOF = true
		end if			
	End Function

    '>>>>>>>> Returns status of the class
	Public Function GetStatus()
		GetStatus = Status
	end function

    '>>>>>>>> Returns Image provided in the RSS/RDF file as a linked image.
	Public Function GetImage()
		if ImageURL <> "" then
			if ImageLink <> "" then GetImage = "<a href=""" & ImageLink & """>"
			GetImage = GetImage & "<img src=""" & ImageURL & """ alt=""" & ImageTitle & """ border=""0"" />"
			if ImageLink <> "" then GetImage = GetImage & "</a>"
		else
			GetImage = ""
		end if
	end function

    '>>>>>>>> Returns the code for the TextInput provided in the RSS/RDF file.
	Public Function GetTextInput()
		if TextInputURL <> "" then
			GetTextInput = "<form method=""post"" action=""" & TextInputURL & """>" & vbCrLf & _
							"<table>" & vbCrLf & _
							"<tr>" & vbCrLf & _
							"	<td>" & TextInputDesc & "</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
							"	<td><input type=""text"" name=""" & TextInputName & """ /></td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
							"	<td><input type=""submit"" value=""" &  TextInputTitle & """ /></td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"</table>" & vbCrLf & _
							"</form>"
		else
			GetTextInput = ""
		end if
	end function

end class
%>
