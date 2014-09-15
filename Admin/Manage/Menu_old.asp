<%
'取得权限信息	
	'QXMC=Session("QXMC")	'类别信息
	YHZL=Session("YHZL")	'用户种类
	YHDL=session("YHDL")	'用户账号
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
	FONT-FAMILY: 宋体;
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
	IF YHZL="管理员" then
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
d = new dTree('d');
d.add(0,-1,'系统功能列表',null,'银监局信息发布系统 系统功能列表','main');
//d.add(1,0,'当前用户',null,'银监局信息发布系统 系统功能列表','main');
//d.add(2,1,'基本信息',"Login.asp",'当前登录帐户的基本信息','main','Images/Skin/AdminList.GIF','Images/Skin/AdminList.GIF');
//d.add(3,1,'注消登录',"Login.asp?Work=LogOut",'','main');
d.add(4,0,'资源管理',null,'','main');
d.add(5,4,'添加资源',"News_Add.asp?Work=AddReco",'添加资源','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
d.add(6,4,'常规资源',"News_List.asp",'常规资源列表','main');
//d.add(7,4,'未审核资源',"News_List.asp?Work=UnChecked",'未审核资源','main');
d.add(7,4,'添加文档新闻',"News/News_Add.asp?Work=AddReco",'添加文档新闻','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
d.add(9,4,'文档新闻列表',"News/News_List.asp",'文档新闻列表','main');
d.add(8,4,'资源回收站',"News_List.asp?Work=Dustbin",'资源回收站','main');
//d.add(40,4,'资源生成',"News_CreateHtml.asp",'资源生成','main');
//d.add(10,4,'附属信息',null,'','main');
//d.add(11,10,'资源特性',"News_Speciality_List.asp",'资源特性','main');
//d.add(12,10,'资源来源',"From_List.asp",'资源来源','main');
//d.add(13,10,'作者列表',"Author_List.asp",'作者列表','main');
//d.add(14,10,'评论列表',"Comment_List.asp",'评论列表','main');
d.add(15,4,'资源分类',null,'','main');
d.add(16,15,'分类列表',"Class_List.asp",'分类列表','main');
//d.add(17,15,'分类模板',"NewsTemplate_List.asp",'模板管理','main');
//d.add(18,15,'添加模板',"NewsTemplate_Mdy.asp?Work=AddReco",'添加模板','main');
//d.add(41,0,'站点更新',null,'','main');
//d.add(42,41,'页面内容替换','InsertSYS_List.asp','','main');
//d.add(43,41,'页面资源更新','UpdatePage.asp','','main');
d.add(19,0,'文件系统',null,'','main');
d.add(20,19,'虚拟目录','FileSystem/View.asp','','main');
d.add(21,0,'系统管理',null,'','main');
d.add(22,21,'参数设置','Sys_Config.asp?Work=MdyFile','','main');
//d.add(28,21,'系统安全',null,'','main');
//d.add(30,28,'添加帐户','Admin_Mdy.asp?Work=AddReco','','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
//d.add(31,28,'帐户列表','Admin_List.asp','','main','Images/Skin/AdminList.GIF','Images/Skin/AdminList.GIF');
//d.add(32,28,'创建角色','AdminRole_Mdy.asp?Work=AddReco','','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
//d.add(33,28,'角色列表','AdminRole_List.asp','','main','Images/Skin/RoleList.GIF','Images/Skin/RoleList.GIF');
//d.add(23,21,'数据库管理',null,'','main');
//d.add(38,23,'数据库统计','DataBase_Statistic.asp','','main','Images/Skin/Report.GIF','Images/Skin/Report.GIF');
//d.add(24,23,'数据库压缩','DataBase_Compact.asp?Work=CompactDB','','main');
//d.add(25,23,'数据库备份','DataBase_Compact.asp?Work=BakDB','','main');
//d.add(26,23,'数据库还原','DataBase_Compact.asp?Work=ReBakDB','','main');
//d.add(27,23,'执行Sql脚本','DataBase_Compact.asp?Work=ExeCuteSql','','main');
//d.add(34,0,'交流回馈',null,'','main');
//d.add(36,34,'官方发布站','http://tsys.ggmmgo.com','','main');
//d.add(37,34,'交流论坛','http://bbs.tsyschina.com','','main','Images/Skin/TalkCenter.GIF','Images/Skin/TalkCenter.GIF');
document.write(d);
//-->
</SCRIPT>
<%Else%>
<SCRIPT LANGUAGE="JavaScript">
<!--
d = new dTree('d');
d.add(0,-1,'系统功能列表',null,'银监局信息发布系统 系统功能列表','main');
//d.add(1,0,'当前用户',null,'银监局信息发布系统 系统功能列表','main');
//d.add(2,1,'基本信息',"Login.asp",'当前登录帐户的基本信息','main','Images/Skin/AdminList.GIF','Images/Skin/AdminList.GIF');
//d.add(3,1,'注消登录',"Login.asp?Work=LogOut",'','main');
d.add(4,0,'资源管理',null,'','main');
d.add(5,4,'添加资源',"News_Add.asp?Work=AddReco",'添加资源','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
d.add(6,4,'常规资源',"News_List.asp",'常规资源列表','main');
d.add(7,4,'添加文档新闻',"News/News_Add.asp?Work=AddReco",'添加文档新闻','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
d.add(9,4,'文档新闻列表',"News/News_List.asp",'文档新闻列表','main');
d.add(8,4,'资源回收站',"News_List.asp?Work=Dustbin",'资源回收站','main');
//d.add(15,4,'资源分类',null,'','main');
//d.add(16,15,'分类列表',"Class_List.asp",'分类列表','main');
d.add(19,0,'文件系统',null,'','main');
d.add(20,19,'虚拟目录','FileSystem/View.asp','','main');
//d.add(21,0,'系统管理',null,'','main');
//d.add(22,21,'参数设置','Sys_Config.asp?Work=MdyFile','','main');
//d.add(28,21,'系统安全',null,'','main');
//d.add(30,28,'添加帐户','Admin_Mdy.asp?Work=AddReco','','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
//d.add(31,28,'帐户列表','Admin_List.asp','','main','Images/Skin/AdminList.GIF','Images/Skin/AdminList.GIF');
//d.add(32,28,'创建角色','AdminRole_Mdy.asp?Work=AddReco','','main','Images/Skin/Add.gif','Images/Skin/Add.gif');
//d.add(33,28,'角色列表','AdminRole_List.asp','','main','Images/Skin/RoleList.GIF','Images/Skin/RoleList.GIF');
document.write(d);
//-->
</SCRIPT>
<%End IF%>
</body>
</html>
