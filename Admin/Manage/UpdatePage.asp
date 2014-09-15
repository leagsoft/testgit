<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<!-- #include file="Include/ClassList_Fun.asp" -->
<!--#include file="Include/CreateFile_Fun.asp" -->
<!--#include file="Include/Tkl_StringClass.asp" -->
<!--#include file="Include/Tkl_TemplateClass.asp" -->
<!--#Include File="Include/OnlineClass.asp" -->
<!--#Include File="Include/UpdateAdminTime.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If

If Not SysAdmin.UpdatePage Then
    Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End If

Call UpdateAdminTime()
%>
<html>
<head>
<title>UpdatePage.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link href="Include/ManageStyle.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" border="0" cellspacing="1" cellpadding="3">
  <tr>
    <td bgcolor="#FFFFCC">[Tsys前台演示站点页面更新]</td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td class="BarTitleBg"><img src="Images/Manage/expand.gif" width="16" height="16">本演示站首页</td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF" class="BarText"> 
      <ul>
        <li>首页全部[<a href="UpdateSite/page01.asp?Work=All">开始更新</a>]</li>
      </ul>
	</td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td class="BarTitleBg"><img src="Images/Manage/expand.gif" width="16" height="16">新闻中心首页</td>
  </tr>
  <tr> 
    <td valign="middle" bgcolor="#FFFFFF" class="BarText"> 
      <ul>
        <li>首页全部[<a href="UpdateSite/page02.asp?Work=All">开始更新</a>]</li>
      </ul>
	</td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td class="BarTitleBg"><img src="Images/Manage/expand.gif" width="16" height="16">图片中心首页</td>
  </tr>
  <tr> 
    <td valign="middle" bgcolor="#FFFFFF" class="BarText"> 
      <ul>
        <li>首页全部[<a href="UpdateSite/page03.asp?Work=All">开始更新</a>]</li>
      </ul>
    </td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td class="BarTitleBg"><img src="Images/Manage/expand.gif" width="16" height="16">下载中心首页</td>
  </tr>
  <tr> 
    <td valign="middle" bgcolor="#FFFFFF" class="BarText"> 
      <ul>
        <li>首页全部[<a href="UpdateSite/page04.asp?Work=All">开始更新</a>]</li>
      </ul>
    </td>
  </tr>
</table>
</body>
</html>