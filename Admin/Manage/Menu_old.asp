<%
'ȡ��Ȩ����Ϣ	
	'QXMC=Session("QXMC")	'�����Ϣ
	YHZL=Session("YHZL")	'�û�����
	YHDL=session("YHDL")	'�û��˺�
%>
<html>
<head>
<title>Menu.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<style type="text/css">
<!--
BODY {
	FONT-FAMILY: ����;
	FONT-SIZE: 9pt;
	SCROLLBAR-HIGHLIGHT-COLOR: buttonface;
	SCROLLBAR-SHADOW-COLOR: buttonface;
	SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;
	SCROLLBAR-TRACK-COLOR: #eeeeee;
	SCROLLBAR-DARKSHADOW-COLOR: buttonshadow
}
-->
</style>
<link rel="StyleSheet" href="Library/DTree/dtree.css" type="text/css">
<SCRIPT src="Library/DTree/dtree.js"></SCRIPT>
</head>
<body bgcolor="#F0F0F0" text="#000000" leftmargin="5" topmargin="5">
<%
	IF YHZL="����Ա" then
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
d = new dTree('d');
d.add(0,-1,'ϵͳ�����б�',null,'�������Ϣ����ϵͳ ϵͳ�����б�','main');
//d.add(1,0,'��ǰ�û�',null,'�������Ϣ����ϵͳ ϵͳ�����б�','main');
//d.add(2,1,'������Ϣ',"Login.asp",'��ǰ��¼�ʻ��Ļ�����Ϣ','main','Images/Skin/AdminList.GIF','Images/Skin/AdminList.GIF');
//d.add(3,1,'ע����¼',"Login.asp?Work=LogOut",'','main');
d.add(4,0,'��Դ����',null,'','main');
d.add(5,4,'�����Դ',"News_Add.asp?Work=AddReco",'�����Դ','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
d.add(6,4,'������Դ',"News_List.asp",'������Դ�б�','main');
//d.add(7,4,'δ�����Դ',"News_List.asp?Work=UnChecked",'δ�����Դ','main');
d.add(7,4,'����ĵ�����',"News/News_Add.asp?Work=AddReco",'����ĵ�����','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
d.add(9,4,'�ĵ������б�',"News/News_List.asp",'�ĵ������б�','main');
d.add(8,4,'��Դ����վ',"News_List.asp?Work=Dustbin",'��Դ����վ','main');
//d.add(40,4,'��Դ����',"News_CreateHtml.asp",'��Դ����','main');
//d.add(10,4,'������Ϣ',null,'','main');
//d.add(11,10,'��Դ����',"News_Speciality_List.asp",'��Դ����','main');
//d.add(12,10,'��Դ��Դ',"From_List.asp",'��Դ��Դ','main');
//d.add(13,10,'�����б�',"Author_List.asp",'�����б�','main');
//d.add(14,10,'�����б�',"Comment_List.asp",'�����б�','main');
d.add(15,4,'��Դ����',null,'','main');
d.add(16,15,'�����б�',"Class_List.asp",'�����б�','main');
//d.add(17,15,'����ģ��',"NewsTemplate_List.asp",'ģ�����','main');
//d.add(18,15,'���ģ��',"NewsTemplate_Mdy.asp?Work=AddReco",'���ģ��','main');
//d.add(41,0,'վ�����',null,'','main');
//d.add(42,41,'ҳ�������滻','InsertSYS_List.asp','','main');
//d.add(43,41,'ҳ����Դ����','UpdatePage.asp','','main');
d.add(19,0,'�ļ�ϵͳ',null,'','main');
d.add(20,19,'����Ŀ¼','FileSystem/View.asp','','main');
d.add(21,0,'ϵͳ����',null,'','main');
d.add(22,21,'��������','Sys_Config.asp?Work=MdyFile','','main');
//d.add(28,21,'ϵͳ��ȫ',null,'','main');
//d.add(30,28,'����ʻ�','Admin_Mdy.asp?Work=AddReco','','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
//d.add(31,28,'�ʻ��б�','Admin_List.asp','','main','Images/Skin/AdminList.GIF','Images/Skin/AdminList.GIF');
//d.add(32,28,'������ɫ','AdminRole_Mdy.asp?Work=AddReco','','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
//d.add(33,28,'��ɫ�б�','AdminRole_List.asp','','main','Images/Skin/RoleList.GIF','Images/Skin/RoleList.GIF');
//d.add(23,21,'���ݿ����',null,'','main');
//d.add(38,23,'���ݿ�ͳ��','DataBase_Statistic.asp','','main','Images/Skin/Report.GIF','Images/Skin/Report.GIF');
//d.add(24,23,'���ݿ�ѹ��','DataBase_Compact.asp?Work=CompactDB','','main');
//d.add(25,23,'���ݿⱸ��','DataBase_Compact.asp?Work=BakDB','','main');
//d.add(26,23,'���ݿ⻹ԭ','DataBase_Compact.asp?Work=ReBakDB','','main');
//d.add(27,23,'ִ��Sql�ű�','DataBase_Compact.asp?Work=ExeCuteSql','','main');
//d.add(34,0,'��������',null,'','main');
//d.add(36,34,'�ٷ�����վ','http://tsys.ggmmgo.com','','main');
//d.add(37,34,'������̳','http://bbs.tsyschina.com','','main','Images/Skin/TalkCenter.GIF','Images/Skin/TalkCenter.GIF');
document.write(d);
//-->
</SCRIPT>
<%Else%>
<SCRIPT LANGUAGE="JavaScript">
<!--
d = new dTree('d');
d.add(0,-1,'ϵͳ�����б�',null,'�������Ϣ����ϵͳ ϵͳ�����б�','main');
//d.add(1,0,'��ǰ�û�',null,'�������Ϣ����ϵͳ ϵͳ�����б�','main');
//d.add(2,1,'������Ϣ',"Login.asp",'��ǰ��¼�ʻ��Ļ�����Ϣ','main','Images/Skin/AdminList.GIF','Images/Skin/AdminList.GIF');
//d.add(3,1,'ע����¼',"Login.asp?Work=LogOut",'','main');
d.add(4,0,'��Դ����',null,'','main');
d.add(5,4,'�����Դ',"News_Add.asp?Work=AddReco",'�����Դ','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
d.add(6,4,'������Դ',"News_List.asp",'������Դ�б�','main');
d.add(7,4,'����ĵ�����',"News/News_Add.asp?Work=AddReco",'����ĵ�����','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
d.add(9,4,'�ĵ������б�',"News/News_List.asp",'�ĵ������б�','main');
d.add(8,4,'��Դ����վ',"News_List.asp?Work=Dustbin",'��Դ����վ','main');
//d.add(15,4,'��Դ����',null,'','main');
//d.add(16,15,'�����б�',"Class_List.asp",'�����б�','main');
d.add(19,0,'�ļ�ϵͳ',null,'','main');
d.add(20,19,'����Ŀ¼','FileSystem/View.asp','','main');
//d.add(21,0,'ϵͳ����',null,'','main');
//d.add(22,21,'��������','Sys_Config.asp?Work=MdyFile','','main');
//d.add(28,21,'ϵͳ��ȫ',null,'','main');
//d.add(30,28,'����ʻ�','Admin_Mdy.asp?Work=AddReco','','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
//d.add(31,28,'�ʻ��б�','Admin_List.asp','','main','Images/Skin/AdminList.GIF','Images/Skin/AdminList.GIF');
//d.add(32,28,'������ɫ','AdminRole_Mdy.asp?Work=AddReco','','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
//d.add(33,28,'��ɫ�б�','AdminRole_List.asp','','main','Images/Skin/RoleList.GIF','Images/Skin/RoleList.GIF');
document.write(d);
//-->
</SCRIPT>
<%End IF%>
</body>
</html>
