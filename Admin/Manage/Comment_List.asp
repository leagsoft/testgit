<!--#include file="../Comment/conn.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#include file="Include/Tkl_StringClass.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="Include/OnlineClass.asp" -->
<!--#Include File="Include/UpdateAdminTime.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not CBool(SysAdmin.Logined) Then
'    Response.Redirect("Login.asp")
'End If

Call UpdateAdminTime()

Dim FunClass
Set FunClass=New Tkl_StringClass
%>
<html>
<head>
<title>Comment_List.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>
<body bgcolor="#FFFFFF">
<%
Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
Dim Sql
    Sql=ExeSql()
    Rs.Open Sql,Conn,1,1
Dim CurrentPage
    If Request("CurrentPage")="" Then
        CurrentPage=1
    Else
        CurrentPage=Request("CurrentPage")
    End If    
    Rs.PageSize=Def_Comment_PageSize
    If Not(Rs.Eof And Rs.Bof) Then
        Rs.AbsolutePage=CurrentPage
    End If
Dim sKey,WorkType
%>
<FORM METHOD=POST ACTION="Comment_Mdy.asp?Work=DelReco" name="DelForm">
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td width="11%" height="24" align="center" class="BarTitleBg">记录ID</td>
    <td width="15%" class="BarTitleBg">发布人</td>
    <td width="15%" class="BarTitleBg">Email</td>
    <td width="31%" class="BarTitleBg">内容</td>
    <td width="14%" class="BarTitleBg">发表时间</td>
    <td width="14%" align="center" class="BarTitleBg">操作 
    </td>
  </tr>
<%
  Dim I
  For I=1 To Rs.PageSize
      If Rs.Eof Then
        Exit For
    End If
%>
  <tr> 
    <td width="11%" height="24" class="BarTitle"><strong><%=Rs("Id")%></strong></td>
    <td width="15%" bgcolor="#FFFFFF"><%=FunClass.HTMLEncode(Rs("title"))%></td>
    <td width="15%" bgcolor="#FFFFFF"><%=FunClass.HTMLEncode(Rs("Email"))%></td>
    <td width="31%" bgcolor="#FFFFFF"><span Title="<%=FunClass.HTMLEncode(Rs("Content"))%>"><%=FunClass.CutStr(FunClass.HTMLEncode(Rs("Content")),40)%></span></td>
    <td width="14%" bgcolor="#FFFFFF"><%=FormatDateTime(Rs("AddTime"),1)%></td>
    <td width="14%" align="center" bgcolor="#FFFFFF"> 
        <INPUT TYPE="checkbox" NAME="ItemChkBox" value="<%=Rs("Id")%>">
    </td>
  </tr>
<%
      Rs.MoveNext
  Next
%>
<%If Rs.Eof And Rs.Bof Then%>
  <tr> 
    <td height="24" bgcolor="#f6f6f6" width="11%" colspan="6" align="center">暂无相关记录</td>
  </tr>
<%End If%>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0">
    <tr> 
      <td align="right"> 
        <script src="Include/Tkl_PageList.js"></script>
        <script>Tkl_PageListBar(<%=Rs.PageCount%>,<%=CurrentPage%>,"Work=<%=Request("Work")%>&sType=<%=Request("sType")%>&sKey=<%=Request("sKey")%>")</script>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="1" cellpadding="2">
    <tr> 
      <td align="right"> 
        <input name="Submit5" type="submit" class="button01-out" value="删  除">
      </td>
    </tr>
  </table>
</FORM>  
<%
Rs.Close
Set Rs=Nothing
%>
</body>
</html>
<%
Function ExeSql()
    Dim tSql
    Select Case Request("Work")
        Case "Search"
            tSql="SELECT * From VisitorComment Where "&Request("sType")&" Like '%"&Request("sKey")&"%' Order By Id DESC"
        Case "ByNews"
            tSql="SELECT * From VisitorComment Where "&Request("sType")&"="&Request("sKey")&" Order By Id DESC"
        Case Else
            tSql="SELECT * From VisitorComment Order By Id DESC"
    End Select
    ExeSql=tSql
End Function
%>