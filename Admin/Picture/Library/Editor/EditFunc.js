//Constants.
var bLoad=false
var bPureText=true
var bodyStyle="<BODY MONOSPACE STYLE=\"font:9pt arial,sans-serif\">"
var bSendAsText=false

public_description=new Editor

function Editor() {
  this.put_html=SetHtml;
  this.get_html=GetHtml;
  this.put_text=SetText;
  this.get_text=GetText;
  this.CompFocus=GetCompFocus;
}

function GetCompFocus() {
  HtmlEditor.focus();
}

function GetText() {
  return HtmlEditor.document.body.innerText;
}

function SetText(text) {
  text = text.replace(/\n/g, "<br>")
  HtmlEditor.document.body.innerHTML=text;
}
 
function GetHtml() {
  if (bSendAsText) 
    return HtmlEditor.document.body.innerText;
  else {
    cleanHtml();
    return HtmlEditor.document.body.innerHTML;
  }
}

function SetHtml(sHtml) {
  if (bSendAsText) HtmlEditor.document.body.innerText=sHtml;
  else HtmlEditor.document.body.innerHTML=sHtml;
}


SEP_PADDING = 5
HANDLE_PADDING = 7

var yToolbars = new Array(); 

var bInitialized = false;
function document.onreadystatechange() {
  if (bInitialized) return;
  bInitialized = true;

  var i, s, curr;

  // Find all the toolbars and initialize them.
  for (i=0; i<document.body.all.length; i++) {
    curr=document.body.all[i];
    if (curr.className == "yToolbar") {
      if (! InitTB(curr)) {
        alert("工具栏：" + curr.id + "初始化失败，请联络系统管理员。");
      }
      yToolbars[yToolbars.length] = curr;
    }
  }

  //Lay out the page, set handler.
  DoLayout();
  window.onresize = DoLayout;

  HtmlEditor.document.open()
  HtmlEditor.document.write("<BODY MONOSPACE STYLE=\"font:10pt arial,sans-serif\"></body>");
  HtmlEditor.document.close()
  HtmlEditor.document.designMode="On"
	loadContent();
}

// Initialize a toolbar button
function InitBtn(btn) {
  btn.onmouseover = BtnMouseOver;
  btn.onmouseout = BtnMouseOut;
  btn.onmousedown = BtnMouseDown;
  btn.onmouseup = BtnMouseUp;
  btn.ondragstart = YCancelEvent;
  btn.onselectstart = YCancelEvent;
  btn.onselect = YCancelEvent;
  btn.YUSERONCLICK = btn.onclick;
  btn.onclick = YCancelEvent;
  btn.bInitialized = true;
  return true;
}

//Initialize a toolbar. 
function InitTB(y) {
  // Set initial size of toolbar to that of the handle
  y.TBWidth = 0;
    
  // Populate the toolbar with its contents
  if (! PopulateTB(y)) return false;
  
  // Set the toolbar width and put in the handle
  y.style.posWidth = y.TBWidth;
  
  return true;
}


// Hander that simply cancels an event
function YCancelEvent() {
  event.returnValue=false;
  event.cancelBubble=true;
  return false;
}

// Toolbar button onmouseover handler
function BtnMouseOver() {
  if (event.srcElement.tagName != "IMG") return false;
  var image = event.srcElement;
  var element = image.parentElement;
  
  // Change button look based on current state of image.
  if (image.className == "Ico") element.className = "BtnMouseOverUp";
  else if (image.className == "IcoDown") element.className = "BtnMouseOverDown";

  event.cancelBubble = true;
}

// Toolbar button onmouseout handler
function BtnMouseOut() {
  if (event.srcElement.tagName != "IMG") {
    event.cancelBubble = true;
    return false;
  }

  var image = event.srcElement;
  var element = image.parentElement;
  yRaisedElement = null;
  
  element.className = "Btn";
  image.className = "Ico";

  event.cancelBubble = true;
}

// Toolbar button onmousedown handler
function BtnMouseDown() {
  if (event.srcElement.tagName != "IMG") {
    event.cancelBubble = true;
    event.returnValue=false;
    return false;
  }

  var image = event.srcElement;
  var element = image.parentElement;

  element.className = "BtnMouseOverDown";
  image.className = "IcoDown";

  event.cancelBubble = true;
  event.returnValue=false;
  return false;
}

// Toolbar button onmouseup handler
function BtnMouseUp() {
  if (event.srcElement.tagName != "IMG") {
    event.cancelBubble = true;
    return false;
  }

  var image = event.srcElement;
  var element = image.parentElement;

  if (element.YUSERONCLICK) eval(element.YUSERONCLICK + "anonymous()");

  element.className = "BtnMouseOverUp";
  image.className = "Ico";

  event.cancelBubble = true;
  return false;
}

// Populate a toolbar with the elements within it
function PopulateTB(y) {
  var i, elements, element;

  // Iterate through all the top-level elements in the toolbar
  elements = y.children;
  for (i=0; i<elements.length; i++) {
    element = elements[i];
    if (element.tagName == "SCRIPT" || element.tagName == "!") continue;
    
    switch (element.className) {
    case "Btn":
      if (element.bInitialized == null) {
	if (! InitBtn(element)) {
	  alert("Problem initializing:" + element.id);
	  return false;
	}
      }
      
      element.style.posLeft = y.TBWidth;
      y.TBWidth += element.offsetWidth + 1;
      break;
      
    case "TBGen":
      element.style.posLeft = y.TBWidth;
      y.TBWidth += element.offsetWidth + 1;
      break;
      
    case "TBSep":
      element.style.posLeft = y.TBWidth + 2;
      y.TBWidth += SEP_PADDING;
      break;
      
    case "TBHandle":
      element.style.posLeft = 2;
      y.TBWidth += element.offsetWidth + HANDLE_PADDING;
      break;
      
    default:
      alert("类出错：" + element.className + " 在元素 " + element.id + " <" + element.tagName + ">");
      return false;
    }
  }

  y.TBWidth += 1;
  return true;
}

function DebugObject(obj) {
  var msg = "";
  for (var i in TB) {
    ans=prompt(i+"="+TB[i]+"\n");
    if (! ans) break;
  }
}

// Lay out the docked toolbars
function LayoutTBs() {
  NumTBs = yToolbars.length;

  // If no toolbars we're outta here
  if (NumTBs == 0) return;

  //Get the total size of a TBline.
  var i;
  var ScrWid = (document.body.offsetWidth) - 70	;
  var TotalLen = ScrWid;
  for (i = 0 ; i < NumTBs ; i++) {
    TB = yToolbars[i];
    if (TB.TBWidth > TotalLen) TotalLen = TB.TBWidth;
  }

  var PrevTB;
  var LastStart = 0;
  var RelTop = 0;
  var LastWid, CurrWid;

  //Set up the first toolbar.
  var TB = yToolbars[0];
  TB.style.posTop = 0;
  TB.style.posLeft = 0;

  //Lay out the other toolbars.
  var Start = TB.TBWidth;
  for (i = 1 ; i < yToolbars.length ; i++) {
    PrevTB = TB;
    TB = yToolbars[i];
    CurrWid = TB.TBWidth;

    if ((Start + CurrWid) > ScrWid) { 
      //TB needs to go on next line.
      Start = 0;
      LastWid = TotalLen - LastStart;
    } 
    else { 
      //Ok on this line.
      LastWid = PrevTB.TBWidth;
      //RelTop -= TB.style.posHeight;
      RelTop -= TB.offsetHeight;
    }
      
    //Set TB position and LastTB width.
    TB.style.posTop = RelTop;
    TB.style.posLeft = Start;
    PrevTB.style.width = LastWid;

    //Increment counters.
    LastStart = Start;
    Start += CurrWid;
  } 

  //Set width of last toolbar.
  TB.style.width = TotalLen - LastStart;
  
  //Move everything after the toolbars up the appropriate amount.
  i--;
  TB = yToolbars[i];
  var TBInd = TB.sourceIndex;
  var A = TB.document.all;
  var item;
  for (i in A) {
    item = A.item(i);
    if (! item) continue;
    if (! item.style) continue;
    if (item.sourceIndex <= TBInd) continue;
    if (item.style.position == "absolute") continue;
    item.style.posTop = RelTop;
  }
}

//Lays out the page.
function DoLayout() {
  LayoutTBs();
}

// Check if toolbar is being used when in text mode
function validateMode() {
  if (! bSendAsText) return true;
  alert("编辑源码状态下不能使用任何设计工具，请切换到设计状态下进行。");
  HtmlEditor.focus();
  return false;
}

//Formats text in HtmlEditor.
function format(what,opt) {
  if (!validateMode()) return;
  if (opt=="removeFormat") {
    what=opt;
    opt=null;
  }

  if (opt==null) HtmlEditor.document.execCommand(what);
  else HtmlEditor.document.execCommand(what,"",opt);
  
  bPureText = false;
  HtmlEditor.focus();
}

//Switches between text and html mode.
function setMode(newMode) {
  bSendAsText = newMode;
  var cont;
  if (bSendAsText) {
    cleanHtml();
    cleanHtml();

    cont=HtmlEditor.document.body.innerHTML;
    HtmlEditor.document.body.innerText=cont;
  } else {
    cont=HtmlEditor.document.body.innerText;
    HtmlEditor.document.body.innerHTML=cont;
  }
  
  HtmlEditor.focus();
}

//Finds and returns an element.
function getEl(sTag,start) {
  while ((start!=null) && (start.tagName!=sTag)) start = start.parentElement;
  return start;
}

function createLink() {
  if (!validateMode()) return;
  
  var isA = getEl("A",HtmlEditor.document.selection.createRange().parentElement());
  var str=prompt("请输入要链接的URL(如：http://www.gotoe.com)：", isA ? isA.href : "http:\/\/");
  
  if ((str!=null) && (str!="http://")) {
    if (HtmlEditor.document.selection.type=="None") {
      var sel=HtmlEditor.document.selection.createRange();
      sel.pasteHTML("<A HREF=\""+str+"\">"+str+"</A> ");
      sel.select();
    }
    else format("CreateLink",str);
  }
  else HtmlEditor.focus();
}

//Sets the text color.
function foreColor() {
  if (! validateMode()) return;
  var arr = showModalDialog("/Oa/Editor/Selcolor.html", "", "font-family:Verdana; font-size:12; dialogWidth:30em; dialogHeight:35em");
  if (arr != null) format('forecolor', arr);
  else HtmlEditor.focus();
}

//Sets the background color.
function backColor() {
  if (!validateMode()) return;
  var arr = showModalDialog("/Oa/Editor/Selcolor.html", "", "font-family:Verdana; font-size:12; dialogWidth:30em; dialogHeight:35em");
  if (arr != null) format('backcolor', arr);
  else HtmlEditor.focus()
}

function cleanHtml() {
  var fonts = HtmlEditor.document.body.all.tags("FONT");
  var curr;
  for (var i = fonts.length - 1; i >= 0; i--) {
    curr = fonts[i];
    if (curr.style.backgroundColor == "#ffffff") curr.outerHTML = curr.innerHTML;
  }
}

function getPureHtml() {
  var str = "";
  var paras = HtmlEditor.document.body.all.tags("P");
  if (paras.length > 0) {
    for (var i=paras.length-1; i >= 0; i--) str = paras[i].innerHTML + "\n" + str;
  } else {
    str = HtmlEditor.document.body.innerHTML;
  }
  return str;
}

function opens(theURL) { 
if (!	validateMode())	return;
window.open(theURL,"","top=200,left=200,toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,width=300,height=110");
}

function fortable()
{
  if (!	validateMode())	return;
  var arr = showModalDialog("/Oa/Editor/Table.html", "", "dialogWidth:13.5em; dialogHeight:10em; status:0");
  
  if (arr != null){
  var ss;
  var rowcolor
  ss=arr.split("*")
  row=ss[0];
  col=ss[1];
  var string;
  string="<table width=100% cellspacing=1 cellpadding=5 bgcolor=#000000>";
  for(i=1;i<=row;i++){
  string=string+"<tr bgcolor=#ffffff height=22>";
  for(j=1;j<=col;j++){
  string=string+"<td></td>";
  }
  string=string+"</tr>";
  }
  string=string+"</table>";
  format('InsertImage');
  content=HtmlEditor.document.body.innerHTML;
  content=content.replace("<IMG>",string);
  HtmlEditor.document.body.innerHTML=content;
  }
  else HtmlEditor.focus();
}

function foreColor()
{
  if (!	validateMode())	return;
  var arr = showModalDialog("/Oa/Editor/Selcolor.html", "", "dialogWidth:18.5em; dialogHeight:19.5em; status:0");
  if (arr != null) format('forecolor', arr);
  else HtmlEditor.focus();
}

function forebr(){
  if (!	validateMode())	return;
  string="<br>";
  format('InsertImage');
  content=HtmlEditor.document.body.innerHTML;
  content=content.replace("<IMG>",string);
  HtmlEditor.document.body.innerHTML=content;
}

// Local Variables:
// c-basic-offset: 2
// End:
  