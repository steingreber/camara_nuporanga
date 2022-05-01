<!-- based on insimage.dlg -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD W3 HTML 3.2//EN">
<HTML  id=dlgImage STYLE="width: 432px; height: 224px; ">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="MSThemeCompatible" content="Yes">
<TITLE>Insert Image</TITLE>
<style>
  html, body, button, div, input, select, fieldset { font-family: MS Shell Dlg; font-size: 8pt; position: absolute; };
</style>
<SCRIPT defer>
function _CloseOnEsc() {
  if (event.keyCode == 27) { window.close(); return; }
}

 document.body.onkeypress = _CloseOnEsc;

   // width to resize large images to
var maxWidth=100;
  // height to resize large images to
var maxHeight=100;
  // valid file types
var fileTypes=["bmp","gif","png","jpg","jpeg"];
  // the id of the preview image tag
var outImage="previewField";
  // what to display when the image is not valid
var defaultPic="spacer.gif";

/***** DO NOT EDIT BELOW *****/

function preview(what){
  var source=what.value;
  var ext=source.substring(source.lastIndexOf(".")+1,source.length).toLowerCase();
  for (var i=0; i<fileTypes.length; i++) if (fileTypes[i]==ext) break;
  globalPic=new Image();
  if (i<fileTypes.length) globalPic.src=source;
  else {
    globalPic.src=defaultPic;
    alert("THAT IS NOT A VALID IMAGE\nPlease load an image with an extention of one of the following:\n\n"+fileTypes.join(", "));
  }
  setTimeout("applyChanges()",200);
}
var globalPic;
function applyChanges(){
  var field=document.getElementById(outImage);
  var x=parseInt(globalPic.width);
  var y=parseInt(globalPic.height);
  if (x>maxWidth) {
    y*=maxWidth/x;
    x=maxWidth;
  }
  if (y>maxHeight) {
    x*=maxHeight/y;
    y=maxHeight;
  }
  field.style.display=(x<1 || y<1)?"none":"";
  field.src=globalPic.src;
  field.width=x;
  field.height=y;
}

</SCRIPT>
</HEAD>
<BODY id=bdy style="background: threedface; color: windowtext;" scroll=no>

	<form action="insert_image_2.asp?acao=1" method="post" enctype="multipart/form-data" name="frmInc" id="frmInc">
<FIELDSET id=fldSpacing style="left: 2em; top: 2em; width: 25em; height: 5.5em;">
<LEGEND id=lgdSpacing>Selecione a imagem em seu computador.</LEGEND>
</FIELDSET>

<input type="file" name="cp_a04foto" id="picField" onchange="preview(this)" style="left: 3.36em; top: 3.8596em; width: 21.72em; height: 2.1294em; ime-mode: disabled;">

<FIELDSET id=fldSpacing style="left: 2em; top: 8em; width: 25em; height: 9em;">
<LEGEND id=lgdSpacing>Preview da imagem...</LEGEND><br />
<img alt="Graphic will preview here" id="previewField" src="spacer.gif">
</FIELDSET>
<BUTTON ID=btnOK style="left: 27.36em; top: 2.4647em; width: 7em; height: 2.2em; " type=submit tabIndex=40>OK</BUTTON>
<BUTTON ID=btnCancel style="left: 27.36em; top: 5.0504em; width: 7em; height: 2.2em; " type=reset tabIndex=45 onClick="window.close();">Cancelar</BUTTON>
	</form>

<%
Dim acao, sNome
acao = Request.QueryString("acao")

If acao = "1" Then

	Response.Expires = 0
	byteCount = Request.TotalBytes
	RequestBin = Request.BinaryRead(byteCount)
	Set UploadRequest = CreateObject("Scripting.Dictionary")
	BuildUploadRequest RequestBin
	
	va_a04foto = Trim(UploadRequest.Item("cp_a04foto").Item("Value"))
	
	If va_a04foto <> "" Then
	ContentType0 = UploadRequest.Item("cp_a04foto").Item("ContentType")
	filepathname0 = UploadRequest.Item("cp_a04foto").Item("FileName")
	FileName0 = Right(filepathname0, Len(filepathname0) - InStrRev(filepathname0, "\"))
	Value0 = UploadRequest.Item("cp_a04foto").Item("Value")
	Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
	numero0 = instrrev(Request.servervariables("Path_Info"), "/")
	Set MyFile0 = ScriptObject.CreateTextFile(Server.mappath("..\fotos") & "\" & FileName0)
	For i = 1 To LenB(Value0)
	MyFile0.Write Chr(AscB(MidB(Value0, i, 1)))
	Next
	MyFile0.Close
	
	va_a04foto = FileName0
	End If
%>
<script>
parent.opener.document.imagem.img23.value = 'fotos/<%= va_a04foto %>';
window.close();
</script>
<%

End If

Sub BuildUploadRequest(RequestBin)
PosBeg = 1
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(13)))
boundary = MidB(RequestBin, PosBeg, PosEnd - PosBeg)
BoundaryPos = InStrB(1, RequestBin, boundary)
Do Until (BoundaryPos = InStrB(RequestBin, boundary & getByteString("--")))
Dim UploadControl
Set UploadControl = CreateObject("Scripting.Dictionary")
Pos = InStrB(BoundaryPos, RequestBin, getByteString("Content-Disposition"))
Pos = InStrB(Pos, RequestBin, getByteString("name="))
PosBeg = Pos + 6
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(34)))
Name = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
PosFile = InStrB(BoundaryPos, RequestBin, getByteString("filename="))
PosBound = InStrB(PosEnd, RequestBin, boundary)
If PosFile <> 0 And (PosFile < PosBound) Then
PosBeg = PosFile + 10
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(34)))
FileName = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
UploadControl.Add "FileName", FileName
Pos = InStrB(PosEnd, RequestBin, getByteString("Content-Type:"))
PosBeg = Pos + 14
PosEnd = InStrB(PosBeg, RequestBin, getByteString(Chr(13)))
ContentType = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
UploadControl.Add "ContentType", ContentType
PosBeg = PosEnd + 4
PosEnd = InStrB(PosBeg, RequestBin, boundary) - 2
Value = MidB(RequestBin, PosBeg, PosEnd - PosBeg)
Else
Pos = InStrB(Pos, RequestBin, getByteString(Chr(13)))
PosBeg = Pos + 4
PosEnd = InStrB(PosBeg, RequestBin, boundary) - 2
Value = getString(MidB(RequestBin, PosBeg, PosEnd - PosBeg))
End If
UploadControl.Add "Value", Value
UploadRequest.Add Name, UploadControl
BoundaryPos = InStrB(BoundaryPos + LenB(boundary), RequestBin, boundary)
Loop
End Sub
'===============================================
Function getByteString(StringStr)
For i = 1 To Len(StringStr)
Char = Mid(StringStr, i, 1)
getByteString = getByteString & ChrB(AscB(Char))
Next
End Function

Function getString(StringBin)
getString = ""
For intCount = 1 To LenB(StringBin)
getString = getString & Chr(AscB(MidB(StringBin, intCount, 1)))
Next
End Function
'===============================================
%>
</body>
</html>
