//初始化
function initToolTip()
{
	document.write("<div id=\"tkl_ToolTip\" style=\"Z-INDEX: 1000;visibility:hidden;POSITION: absolute;left: 0px; top: 0px;width:1px;\">")
	document.write("  <table width=\"100%\" border=\"0\" cellspacing=\"1\" cellpadding=\"2\" bgcolor=\"#000000\">")
	document.write("    <tr>")
	document.write("      <td bgcolor=\"#FFFFCC\" id=\"tkl_ToolTip_Content\" nowrap></td>")
	document.write("    </tr>")
	document.write("  </table>")
	document.write("</div>")
}
//函数：显示标注
//参数：内容,event.srcElement
var Tkl_Tooltip_EventObj=null
var Tkl_Tooltip_Show=false
function showToolTip(str,obj)
{
	Tkl_Tooltip_EventObj=obj
	Tkl_Tooltip_Show=true
	setTimeout("showToolTip2('"+str+"')",1000)
}
function showToolTip2(str) 
{
	if(!Tkl_Tooltip_Show){return}
	var obj=Tkl_Tooltip_EventObj
	if(str==""){return}
	var t=obj.offsetTop;
	var l=obj.offsetLeft;
	var h=obj.offsetHeight;
	while(obj=obj.offsetParent){
		t+=obj.offsetTop;
		l+=obj.offsetLeft;
	}
	tkl_ToolTip.style.top=t+h+5
	tkl_ToolTip.style.left=l
	tkl_ToolTip.style.visibility="visible"
	tkl_ToolTip_Content.innerHTML=str
}
//函数：隐藏标注
function hiddenToolTip()
{
	Tkl_Tooltip_Show=false
	tkl_ToolTip.style.visibility="hidden"
	tkl_ToolTip_Content.innerHTML=""
}

initToolTip()