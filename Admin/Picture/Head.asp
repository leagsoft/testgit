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
<script src="Library/htmlarea/init_htmlarea.js"></script>
<title>广东省银监局宣传园地管理系统</title>
</head>
<body topmargin="0" leftmargin="0" marginwidth = "0" marginheight = "0" bgcolor="white">
<!--菜单开始-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="100%" bgcolor="#FFDFDF" height="30" align=center>
<a href="index.asp" title="宣传园地管理。">宣传园地管理</a>
    </td>
  </tr>
  <tr>
    <td width="100%" bgcolor="#990000" height="1"></td>
  </tr>
</table>
<!--菜单结束-->
<br>
<br>
