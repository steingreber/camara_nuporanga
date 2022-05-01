function fotoZoom(img){
  foto1= new Image();
  foto1.src=(img);
  Controlla(img);
}
function Controlla(img){
  if((foto1.width!=0)&&(foto1.height!=0)){
    viewFoto(img);
  }
  else{
    funzione="Controlla('"+img+"')";
    intervallo=setTimeout(funzione,20);
  }
}
function viewFoto(img){
  largh=foto1.width+20;
  altez=foto1.height+20;
  stringa="width="+largh+",height="+altez;
  finestra=window.open(img,"",stringa);
}

//Valida Login de acesso
function validalogin(theForm)
{
  if (theForm.email.value == "")
  {
    alert("Digite um valor para o campo EMAIL");
    theForm.email.focus();
    return (false);
  }
  if (theForm.senha.value == "")
  {
    alert("Digite um valor para o campo SENHA");
    theForm.senha.focus();
    return (false);
  }
  return (true);
}

//VALIDA CONTATO
function Validator(theForm)
{
  if (theForm.cp1NOME.value == "")
  {
    alert("O campo NOME não pode ser vazio!");
    theForm.cp1NOME.focus();
    return (false);
  }
  if (theForm.cp2CIDADE.value == "")
  {
    alert("O campo CIDADE não pode ser vazio!");
    theForm.cp2CIDADE.focus();
    return (false);
  }
  if (theForm.cp3ESTADO.value == "")
  {
    alert("O campo ESTADO não pode ser vazio!");
    theForm.cp3ESTADO.focus();
    return (false);
  }
  if (theForm.cp4DDD.value == "")
  {
    alert("O campo DDD não pode ser vazio!");
    theForm.cp4DDD.focus();
    return (false);
  }
  if (theForm.cp5FONE.value == "")
  {
    alert("O campo FONE não pode ser vazio!");
    theForm.cp5FONE.focus();
    return (false);
  }
  if (theForm.cp6EMAIL.value == "")
  {
    alert("O campo EMAIL não pode ser vazio!");
    theForm.cp6EMAIL.focus();
    return (false);
  }
  return (true);
}

//Valida cadastro de acesso
function validacadastro(theForm)
{
  if (theForm.fu30username.value == "")
  {
    alert("Digite um valor para o campo EMAIL");
    theForm.fu30username.focus();
    return (false);
  }
  if (theForm.fu31senha.value == "")
  {
    alert("Digite um valor para o campo SENHA");
    theForm.fu31senha.focus();
    return (false);
  }
  if (theForm.fu01Nome.value == "")
  {
    alert("Digite um valor para o campo NOME ");
    theForm.fu01Nome.focus();
    return (false);
  }
  if (theForm.fu04FoneContato.value == "")
  {
    alert("Digite um valor para o campo FONE CONTATO");
    theForm.fu04FoneContato.focus();
    return (false);
  }
  if (theForm.fu05EstadoCivil.value == "")
  {
    alert("Selecione um valor para o campo ESTADO CIVIL");
    theForm.fu05EstadoCivil.focus();
    return (false);
  }
  if (theForm.fu06Sexo.value == "")
  {
    alert("Selecione um valor para o campo SEXO");
    theForm.fu06Sexo.focus();
    return (false);
  }
  if (theForm.fu07Identidade.value == "")
  {
    alert("Digite um valor para o campo RG");
    theForm.fu07Identidade.focus();
    return (false);
  }
  if (theForm.fu06CPF.value == "")
  {
    alert("Digite um valor para o campo CPF");
    theForm.fu06CPF.focus();
    return (false);
  }
  if (theForm.fu14EndRes.value == "")
  {
    alert("Digite um valor para o campo ENDEREÇO");
    theForm.fu14EndRes.focus();
    return (false);
  }
  if (theForm.fu15Numero.value == "")
  {
    alert("Digite um valor para o campo NÚMERO");
    theForm.fu15Numero.focus();
    return (false);
  }  
  if (theForm.fu19Cidade.value == "")
  {
    alert("Digite um valor para o campo CIDADE");
    theForm.fu19Cidade.focus();
    return (false);
  }
  if (theForm.fu20Estado.value == "")
  {
    alert("Digite um valor para o campo ESTADO");
    theForm.fu20Estado.focus();
    return (false);
  }
  if (theForm.fu18CEP.value == "")
  {
    alert("Digite um valor para o campo CEP");
    theForm.fu18CEP.focus();
    return (false);
  }
  if (theForm.fu02Data_Nascto.value == "")
  {
    alert("Digite um valor para o campo DATA NASCTO:");
    theForm.fu02Data_Nascto.focus();
    return (false);
  }
  return (true);
}

// Desativa a visualização do código ///////////////////////////////////
function clickIE() {if (document.all) {(message);return false;}}
function clickNS(e) {if 
(document.layers||(document.getElementById&&!document.all)) {
if (e.which==2||e.which==3) {(message);return false;}}}
if (document.layers) 
{document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;}
else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;}

document.oncontextmenu=new Function("return false")

//script pro menu principal - Muda fundo da TD
function mOvr(src,clrOver) {
 if (!src.contains(event.fromElement)) {
	 src.style.cursor = 'hand';
	 src.bgColor = clrOver;
	}
 }
 function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	 src.style.cursor = 'default';
	 src.bgColor = clrIn;
	}
 }
 function mClk(src) {
if(event.srcElement.tagName=='TD'){
	src.children.tags('A')[0].click();
}
}
//-->

function goto_byselect(sel, targetstr)
{
  var index = sel.selectedIndex;
  if (sel.options[index].value != '') {
     if (targetstr == 'blank') {
       window.open(sel.options[index].value, 'win1');
     } else {
       var frameobj;
       if (targetstr == '') targetstr = 'self';
       if ((frameobj = eval(targetstr)) != null)
         frameobj.location = sel.options[index].value;
     }
  }
}

function ConteudoFlash(nomefla, altura, largura){
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="'+largura+'" height="'+altura+'" id="c4" align="middle">\n');
document.write('<param name="allowScriptAccess" value="sameDomain" />\n');
document.write('<param name="Menu" value="false">');
document.write('<param name="movie" value="'+nomefla+'" /><param name="quality" value="high" /><param name="wmode" value="transparent" /><param name="bgcolor" value="#ffffff" /><embed src="'+nomefla+'.swf" quality="high" wmode="transparent" bgcolor="#ffffff" width="'+largura+'" height="'+altura+'" name="top_ micro" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />\n');
document.write('</object>\n');
}

function ajax(url,div) 
{ 
document.getElementById(div).innerHTML='<div align="center" class="carregando"><br><img src="imagens/carregando.gif"><br /><br />aguarde...</div>'
mostra=document.getElementById(div);
    req = null; 
    if (window.XMLHttpRequest) { 
        req = new XMLHttpRequest(); 
        req.onreadystatechange = processReqChange; 
        req.open("GET",url,true); 
        req.send(null); 
    } else if (window.ActiveXObject) { 
        req = new ActiveXObject("Microsoft.XMLHTTP"); 
        if (req) {
		
         req.onreadystatechange = processReqChange; 
         req.open("GET",url,true); 
	 
            req.send(); 
        } 
    } 
} 

function processReqChange() 
{
    if (req.readyState == 4) { 
        if (req.status ==200) { 
			mostra.innerHTML = req.responseText;
 
        } else { 
            alert("Houve um problema ao obter os dados:\n" + req.statusText); 
        } 
    } 
} 

function open_window(name, url, left, top, width, height, toolbar, menubar, statusbar, scrollbar, resizable)
{
  toolbar_str = toolbar ? 'yes' : 'no';
  menubar_str = menubar ? 'yes' : 'no';
  statusbar_str = statusbar ? 'yes' : 'no';
  scrollbar_str = scrollbar ? 'yes' : 'no';
  resizable_str = resizable ? 'yes' : 'no';

  cookie_str = document.cookie;
  cookie_str.toString();

  pos_start  = cookie_str.indexOf(name);
  pos_end    = cookie_str.indexOf('=', pos_start);

  cookie_name = cookie_str.substring(pos_start, pos_end);

  pos_start  = cookie_str.indexOf(name);
  pos_start  = cookie_str.indexOf('=', pos_start);
  pos_end    = cookie_str.indexOf(';', pos_start);
  
  if (pos_end <= 0) pos_end = cookie_str.length;
  cookie_val = cookie_str.substring(pos_start + 1, pos_end);
  if (cookie_name == name && cookie_val  == "done")
    return;

  window.open(url, name, 'left='+left+',top='+top+',width='+width+',height='+height+',toolbar='+toolbar_str+',menubar='+menubar_str+',status='+statusbar_str+',scrollbars='+scrollbar_str+',resizable='+resizable_str);
}