<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="Include/OnlineClass.asp" -->
<!--#Include File="Include/UpdateAdminTime.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If

Call UpdateAdminTime()
%>
<html>
<head>
<title>From_List.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>

<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td> 
      <input name="Submit32" type="button" class="button02-out" value="�����Դ" onClick="window.location='From_Mdy.asp?Work=AddReco'" Title="�����Դ"></td>
  </tr>
</table>
<%
Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
Dim Sql
    Sql="SELECT * From FromList Order By upTime DESC"
    Rs.Open Sql,Conn,1,1
Dim CurrentPage
    If Request("CurrentPage")="" Then
        CurrentPage=1
    Else
        CurrentPage=Request("CurrentPage")
    End If    
    Rs.PageSize=20
    If Not(Rs.Eof And Rs.Bof) Then
        Rs.AbsolutePage=CurrentPage
    End If
Dim sKey,WorkType
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr class="BarTitleBg"> 
    <td width="9%" height="15">��¼ID</td>
    <td>��Դ����</td>
    <td>��ʾЧ��</td>
    <td width="17%">����ʱ��</td>
    <td width="12%" align="center">�༭</td>
  </tr>
<%
  Dim I
  For I=1 To Rs.PageSize
      If Rs.Eof Then
        Exit For
    End If
%>
  <tr> 
    <td width="9%" height="24" bgcolor="#FFFFFF" class="BarTitle"><%=Rs("Id")%></td>
    <td width="16%" bgcolor="#FFFFFF"><%=Rs("Title")%></td>
    <td width="46%" bgcolor="#FFFFFF"><%=Rs("Content")%></td>
    <td width="17%" bgcolor="#FFFFFF"><%=FormatDateTime(Rs("upTime"),2)%></td>
    <td width="12%" align="center" bgcolor="#FFFFFF">
<input name="Submit3" type="button" class="button01-out" value="��  ��" onClick="window.location='From_Mdy.asp?id=<%=Rs("Id")%>'"></td>
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
      <script>Tkl_PageListBar(<%=Rs.PageCount%>,<%=CurrentPage%>,"")</script>
    </td>
  </tr>
</table>
<%
Rs.Close
Set Rs=Nothing
%>
</body>
</html>
