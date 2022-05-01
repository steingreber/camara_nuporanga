<HTML><HEAD><TITLE>PageMaster 2.0</TITLE>
<script Language="JavaScript">
// Begin
function Validator(theForm)
{
  if (theForm.username.value == "")
  {
    alert("Usuário!!");
    theForm.username.focus();
    return (false);
  }
  if (theForm.password.value == "")
  {
    alert("Senha!!");
    theForm.password.focus();
    return (false);
  }
  return (true);
}
function placeFocus() {
if (document.forms.length > 0) {
var field = document.forms[0];
for (i = 0; i < field.length; i++) {
if ((field.elements[i].type == "text") || (field.elements[i].type == "textarea") || (field.elements[i].type.toString().charAt(0) == "s")) {
document.forms[0].elements[i].focus();
break;
         }
      }
   }
}
// End
</script>

<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<META content="robots" name="none"></HEAD>
<body bgcolor="#FFFFFF" link="#0038A8" vlink="#666666" alink="#00A3D0" onLoad="placeFocus();">
<DIV align=center>
<TABLE height="70%" cellSpacing=0 cellPadding=0 width="100%" border=0>
  <form action="pm_verifica.asp" method="post" onsubmit="return Validator(this)">
  <INPUT type=hidden value=loginauth name=action>
  <TBODY>
  <TR vAlign=center align=middle>
    <TD><IMG height=27 alt="PageMaster Login" src="images/mc-tab-login.gif" width=300 border=0><BR>
      <TABLE cellSpacing=0 cellPadding=0 width=300 bgColor=#dcdfe1 border=0>
        <TBODY>
        <TR vAlign=center>
          <TD width=7 height=3><IMG height=3 src="images/background-tl.gif" width=7></TD>
          <TD width=286 background="images/background-t.gif"
          height=3><IMG height=3 src="images/background-t.gif" width=286></TD>
          <TD width=7 height=3><IMG height=3 src="images/background-tr.gif" width=7></TD></TR>
        <TR vAlign=center>
          <TD width=7 background=images/background-l.gif>&nbsp;</TD>
          <TD width=286>
            <TABLE cellSpacing=0 cellPadding=5 align=center border=0>
              <TBODY>
              <TR>
                <TD>
                  <TABLE cellSpacing=0 cellPadding=5 align=center border=0>
                    <TBODY>
                    <TR>
                      <TD><FONT face="Arial, Helvetica, sans-serif" size=2>Usuário</FONT></TD>
                      <TD><INPUT size=22 name=username></TD></TR>
                    <TR>
                      <TD><FONT face="Arial, Helvetica, sans-serif" size=2>Senha</FONT></TD>
                      <TD><INPUT type=password size=22 name=password> 
                    </TD></TR></TBODY></TABLE>
                  <DIV align=center><IMG height=5 src="images/clear-pixel.gif" width=5><BR>
                    <p><INPUT type=image height=23 alt=Login width=60 src="images/button-login.gif" border=0></DIV>
                   </TD></TR></TBODY></TABLE></TD>
          <TD width=7 background=images/background-r.gif>&nbsp;</TD></TR>
        <TR vAlign=bottom>
          <TD height=7><IMG height=7 src="images/background-bl.gif" width=7></TD>
          <TD width=286 background=images/background-b.gif 
          height=7><IMG height=7 src="images/background-b.gif" width=286></TD>
          <TD height=7><IMG height=7 src="images/background-br.gif" width=7></TD>
         </TR></TBODY></TABLE></TD></TR></FORM></TBODY></TABLE></DIV></BODY></HTML>

