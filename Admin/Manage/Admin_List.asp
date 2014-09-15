<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="Include/OnlineClass.asp" -->
<!--#Include File="Include/UpdateAdminTime.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not SysAdmin.Logined Then
    Response.Redirect("Login.asp")
End If

Call UpdateAdminTime()
%>
<html>
<head>
<title>Admin_List.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>

<body bgcolor="#FFFFFF">

<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td> <input name="Button2" type="button" class="button02-out" value="添加帐户" onClick="window.location='Admin_Mdy.asp?Work=AddReco'"></td>
  </tr>
</table>
<%
Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
Dim Sql
    Sql="SELECT * From View_AdminInfo Order By Id"
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
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
  <tr bgcolor="#CCCCCC" class="BarTitleBg"> 
    <td width="4%" height="15">&nbsp;</td>
    <td>帐户</td>
    <td>角色</td>
    <td align="center">添加时间</td>
    <td width="16%" align="center">更新时间</td>
    <td width="12%" align="center">编辑</td>
  </tr>
<%
  Dim I
  Dim EPopeDom
  EPopeDom=SysAdmin.ChangeAdminList
  Dim Show
  Show=false
  For I=1 To Rs.PageSize
      If Rs.Eof Then
        Exit For
    End If
    If EPopeDom Then
        Show=true
    Else
        If UCase(Rs("Title"))=UCase(Session("AdminTitle"))  Then
            Show=true
        Else
            Show=false
        End If
    End If
    If Show Then
%>
  <tr bgcolor="#FFFFFF"> 
    <td width="4%" height="27" align="center"><img src="Images/Manage/Admin<%If Not CBool(Rs("Lock")) Then Response.Write("Un") End If%>Lock.gif" width="16" height="16" Title="记录Id:<%=Rs("Id")%>"></td>
    <td width="21%"><%=Rs("Title")&" <font color=""#666666"">("& Rs("NickName") &")</font>"%></td>
    <td width="31%"><%If Rs("RoleTitle")=SysAdmin.AdminRoleTitle Then Response.Write("<strong>"& Rs("RoleTitle") &"</strong>") Else Response.Write(Rs("RoleTitle")) End If%>
    </td>
    <td width="16%" align="center"><%=FormatDateTime(Rs("AddTime"),2)%></td>
    <td width="16%" align="center"><%=FormatDateTime(Rs("upTime"),2)%></td>
    <td width="12%" align="center"><input name="Button" type="button" class="button01-out" onClick="window.location='Admin_Mdy.asp?id=<%=Rs("Id")%>'" value="编  辑">
    </td>
  </tr>
<%
    End If
      Rs.MoveNext
  Next
%>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0">
  <tr> 
    <td align="right"> 
      <script src="Include/Tkl_PageList.js"></script>
      <script>Tkl_PageListBar(<%=Rs.PageCount%>,<%=CurrentPage%>,"")</script>
    </td>
  </tr>
</table>
<%
Rs.Close
Set Rs=Nothing
%>
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display=''}else{HelpTab.style.display='none'}">&nbsp;<img src="Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:none"> 
    <td width="2%">&nbsp;</td>
    <td width="98%"><table width="100%" border="0" cellspacing="3" cellpadding="0">
        <tr> 
          <td><img src="Images/Manage/AdminLock.gif" width="16" height="16"> 
            表示[帐户]已被锁定，无法使用。<br>
            <img src="Images/Manage/AdminUnLock.gif" width="16" height="16"> 
            表示[帐户]未被锁定，可正常使用。</td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
