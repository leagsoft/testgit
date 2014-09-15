<!--#Include File=Include/Config.asp-->
<%
	Session("QXMC")=Trim(Request("QXMC"))		'权限名称
	Session("Purview")=Trim(Request("Purview"))	'用户权限
	Session("Column")=Trim(Request("Column"))	'用户控制的栏目
	'Session("YHZL")=Trim(Request("YHZL"))		'用户种类
%>
<html>
<head>
<title><%=Def_SysTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<style>
body {
    margin: 0px;
    padding: 0px;
    text-align: center;
    border: none;
}
</style>
</head>
<frameset rows="49,*" cols="*" framespacing="5" frameborder="yes" border="5" bordercolor="#CCCCCC">
  <frame src="Top.asp" name="frameTop" frameborder="no" scrolling="NO" noresize bordercolor="#CCCCCC">
  <frameset rows="*" cols="160,*" framespacing="4" frameborder="yes" border="4" bordercolor="#CCCCCC" id="fram1">
    <frame src="Menu.asp" name="menu" frameborder="no" scrolling="auto" bordercolor="#CCCCCC">
    <!--<frame src="Login.asp" name="main" frameborder="no" bordercolor="#CCCCCC">-->
    <frame src="News_list.asp" name="main" frameborder="no" bordercolor="#CCCCCC">
  </frameset>
</frameset>
<noframes> 
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>