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
<title>�㶫ʡ�����ͼƬ����ϵͳ</title>
<!--�����ߣ����ݿ��ӿƼ����޹�˾-->
</head>
<body topmargin="0" leftmargin="0" marginwidth = "0" marginheight = "0" bgcolor="white">
<!--�˵���ʼ-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="100%" bgcolor="#FFDFDF" height="30" align=center>
<!--<a href="Menu.asp" title="���غ�̨����ϵͳ��ҳ">��ҳ</a>&nbsp;-&nbsp;
<a href="DictList.asp" title="ά����վ�еĸ��ַ�����Ϣ"><%If cNav="dict" Then Response.Write "<font color=red>"%>�ֵ����</font></a>&nbsp;-&nbsp;
<a href="UPicture.asp" title="ά����վ��һ������ı�־ͼ"><%If cNav="pic" Then Response.Write "<font color=red>"%>�����־</font></a>&nbsp;-&nbsp;-->
<a href="index.asp" title="ά��ϵͳ�����Ϣ��">ͼƬ����</a>
<!--<a href="ENQUIRYList.asp" title="ά����Ʒ��Ϣ���鿴ѯ�۵���"><%If cNav="ENQUIRY" Then Response.Write "<font color=red>"%>ѯ�̵�����</font></a>&nbsp;-&nbsp;
<a href="LogFileList.asp" title="ά����Ʒ��Ϣ���鿴ѯ�۵���"><%If cNav="Log" Then Response.Write "<font color=red>"%>��־����</font></a>&nbsp;-&nbsp;
<a href="GuestList.asp" title="�鿴�ͻ�������Ϣ��"><%If cNav="guest" Then Response.Write "<font color=red>"%>��������</font></a>&nbsp;-&nbsp;
<a href="NewsList.asp" title="ά����վע���û���"><%If cNav="news" Then Response.Write "<font color=red>"%>��Ϣ����</font></a>&nbsp;-&nbsp;
<a href="UserList.asp" title="ά����վ������Ϣ��"><%If cNav="user" Then Response.Write "<font color=red>"%>�û�����</font></a>&nbsp;-&nbsp;
<a href="Index.asp?cAction=<%= Server.UrlEncode("��  ��")%>" title="�˳���̨����ϵͳ">�˳�</a>-->

    </td>
  </tr>
  <tr>
    <td width="100%" bgcolor="#990000" height="1"></td>
  </tr>
</table>
<!--�˵�����-->
<br>
<br>
