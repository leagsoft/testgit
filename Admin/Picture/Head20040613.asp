<html>
<head>
<script LANGUAGE="JavaScript">
<!--
function GoToList()
{	//alert(document.form.cAction)
	document.GoList.submit()
}
function DoConfirm(cMsg,cUpdate)
{
	if(confirm(cMsg))
	{
		document.UserForm.fcUpdate.value = cUpdate
		document.UserForm.submit()
	}
}
function GoToForm(nRecNum)
{
	document.GoForm.fnStartRecord.value = nRecNum
	document.GoForm.submit()
}
function js_t(htmlurl) {
var
newwin=window.open(htmlurl,"vm","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,width=450,height=500,top=0,left=200");
  return false;
}
function js_g(htmlurl) {
var
newwin=window.open(htmlurl,"message","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=400,top=0,left=200");
  return false;
}
//-->
</script>
<script RUNAT="Server" LANGUAGE="VBScript">
Function DateToString(dDate)
   DateToString = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)+RIGHT("00"+Trim(Hour(dDate)),2)+RIGHT("00"+Trim(Minute(dDate)),2)+RIGHT("00"+Trim(Second(dDate)),2)
End Function

FUNCTION TimeToString(tTime)
    TimeToString = Trim(Hour(tTime)) + ":" + Trim(Minute(tTime)) + ":" + Trim(Second(tTime))
END FUNCTION
</script>
<link rel="stylesheet" href="GreatSoft.css" type="text/css">
<title>广东省银监局图片管理系统</title>
<!--开发者：广州竣视科技有限公司-->
</head>
<body topmargin="0" leftmargin="0" marginwidth = "0" marginheight = "0" bgcolor="white">
<!--菜单开始-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="100%" bgcolor="#FFDFDF" height="30" align=center>
<!--<a href="Menu.asp" title="返回后台管理系统首页">首页</a>&nbsp;-&nbsp;
<a href="DictList.asp" title="维护网站中的各种分类信息"><%If cNav="dict" Then Response.Write "<font color=red>"%>字典管理</font></a>&nbsp;-&nbsp;
<a href="UPicture.asp" title="维护网站中一级分类的标志图"><%If cNav="pic" Then Response.Write "<font color=red>"%>分类标志</font></a>&nbsp;-&nbsp;-->
<a href="index.asp" title="维护系统码表信息。">图片管理</a>
<!--<a href="ENQUIRYList.asp" title="维护产品信息及查看询价单。"><%If cNav="ENQUIRY" Then Response.Write "<font color=red>"%>询盘单管理</font></a>&nbsp;-&nbsp;
<a href="LogFileList.asp" title="维护产品信息及查看询价单。"><%If cNav="Log" Then Response.Write "<font color=red>"%>日志管理</font></a>&nbsp;-&nbsp;
<a href="GuestList.asp" title="查看客户反馈信息。"><%If cNav="guest" Then Response.Write "<font color=red>"%>反馈管理</font></a>&nbsp;-&nbsp;
<a href="NewsList.asp" title="维护网站注册用户。"><%If cNav="news" Then Response.Write "<font color=red>"%>信息管理</font></a>&nbsp;-&nbsp;
<a href="UserList.asp" title="维护网站公布信息。"><%If cNav="user" Then Response.Write "<font color=red>"%>用户管理</font></a>&nbsp;-&nbsp;
<a href="Index.asp?cAction=<%= Server.UrlEncode("退  出")%>" title="退出后台管理系统">退出</a>-->

    </td>
  </tr>
  <tr>
    <td width="100%" bgcolor="#990000" height="1"></td>
  </tr>
</table>
<!--菜单结束-->
<br>
<br>
