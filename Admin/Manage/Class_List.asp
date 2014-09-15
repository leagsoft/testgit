<!--#include file="Include/Conn.asp" -->
<!-- #include file="Include/ClassList_Fun.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="Include/OnlineClass.asp" -->
<!--#Include File="Include/UpdateAdminTime.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not SysAdmin.Logined Then
'    Response.Redirect("Login.asp")
'End If

Call UpdateAdminTime()

Dim Parent
If Request("Parent")="" Then
    Parent=SysAdmin.AdminTopClassId
Else
    Parent=CLng(Request("Parent"))
End If
%>
<html>
<head>
<title>Class_List.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>

<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr>
    <td> 
      <input name="Submit32" type="button" class="button02-out" value="添加分类" onClick="window.location='Class_Mdy.asp?Work=AddReco&Parent=<%=Parent%>'" title="在当前位置添加新分类"></td>
  </tr>
</table>
<%
Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
Dim Sql
    Sql="SELECT * From ClassList Where Parent="&Parent&" Order By OrderNum DESC,upTime DESC"
    Rs.PageSize=20
	Rs.CacheSize=Rs.PageSize
    Rs.Open Sql,Conn,1,1
Dim CurrentPage
    If Request("CurrentPage")="" Then
        CurrentPage=1
    Else
        CurrentPage=Request("CurrentPage")
    End If    
    If Not(Rs.Eof And Rs.Bof) Then
        Rs.AbsolutePage=CurrentPage
    End If
Dim sKey,WorkType
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td height="11" colspan="6" bgcolor="#FFFFFF"> <%=GetClassPath2(SysAdmin.AdminTopClassId,Parent,"")%> 的子分类列表：</td>
  </tr>
  <tr class="BarTitleBg"> 
    <td width="9%" height="11">记录ID</td>
    <td width="22%">资源分类名称</td>
    <!--<td width="22%">生成目录</td>-->
    <td width="18%" align="center">资源列表</td>
    <td width="17%" align="center">更新时间</td>
    <td width="12%" align="center">编辑</td>
  </tr>
  <%
  Dim I
  For I=1 To Rs.PageSize
      If Rs.Eof Then
        Exit For
    End If
%>
  <tr> 
    <td width="9%" height="28" bgcolor="#FFFFFF" class="BarTitle"><%=Rs("Id")%></td>
    <td bgcolor="#FFFFFF"><a href="?Parent=<%=Rs("Id")%>"><%=Rs("title")%></a></td>
    <!--<td bgcolor="#FFFFFF"><a href="Class_Mdy.asp?Work=DirectoryInfo&Id=<%=Rs("Id")%>"><%=Rs("Directory")%></a></td>-->
    <td align="center" bgcolor="#FFFFFF"><a href="News_List.asp?Parent=<%=Parent%>">资源列表</a></td>
    <td width="17%" align="center" bgcolor="#FFFFFF"><%=FormatDateTime(Rs("upTime"),2)%></td>
    <td width="12%" align="center" bgcolor="#FFFFFF"> 
      <input name="Submit3" type="button" class="button01-out" value="编  辑" onclick="window.location='Class_Mdy.asp?id=<%=Rs("Id")%>'">
    </td>
  </tr>
  <%
      Rs.MoveNext
  Next
%>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0">
  <tr> 
    <td align="right"> 
      <script src="Include/Tkl_PageList.js"></script>
      <script>Tkl_PageListBar(<%=Rs.PageCount%>,<%=CurrentPage%>,"Parent=<%=Parent%>")</script>
    </td>
  </tr>
</table>
<%
Rs.Close
Set Rs=Nothing
%>
</body>
</html>
