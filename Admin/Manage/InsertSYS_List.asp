<!--#include file="Include/Conn.asp" -->
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
%>
<html>
<head>
<title>InsertSYS_List.asp</title>
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
      <input name="Submit32" type="button" class="button02-out" value="���Ƕ��" onClick="window.location='InsertSYS_Mdy.asp?Work=AddReco'"> 
      <input name="Submit322" type="button" class="button02-out" value="ִ��Ƕ��" onClick="window.location='InsertSYS_Mdy.asp?Work=InsertSysActive'"></td>
  </tr>
</table>
<%
Dim Work
    Work=Request("Work")
Dim sType
    sType=Replace(Request("sType"),"'","")
    If sType="" Then
        sType="Title"
    End If
Dim sKey
    sKey=Replace(Request("sKey"),"'","")

Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
Dim Sql
    Sql=ExeSql()
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
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr class="BarTitleBg"> 
    <td width="9%" height="15">��¼ID</td>
    <td>����</td>
    <td align="center">����ʱ��</td>
    <td width="20%" align="center">���ʱ��</td>
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
    <td width="34%" bgcolor="#FFFFFF"><%=Rs("Title")%></td>
    <td width="25%" align="center" bgcolor="#FFFFFF"><%=Rs("upTime")%></td>
    <td width="20%" align="center" bgcolor="#FFFFFF"><%=Rs("AddTime")%></td>
    <td width="12%" align="center" bgcolor="#FFFFFF"><input name="Submit3" type="button" class="button01-out" value="��  ��" onClick="window.location='InsertSYS_Mdy.asp?id=<%=Rs("Id")%>'">
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
      <script>Tkl_PageListBar(<%=Rs.PageCount%>,<%=CurrentPage%>,"Work=<%=Work%>&sType=<%=sType%>&sKey=<%=sKey%>")</script>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <form name="form1" method="post" action="?" onsubmit="return chkSearchForm(this)">
    <tr bgcolor="#FFFFFF"> 
      <td width="67%" align="right"><a name="AdvanceSh"></a> 
        <input name="Work" type="hidden" id="Work" value="<%=Work%>">
        ����: 
        <select name="sType" class="Input">
          <option value="Title" <%If sType="Title" Then Response.Write("selected") End If%>>�ꡡ��</option>
          <option value="Content" <%If sType="Content" Then Response.Write("selected") End If%>>��������</option>
        </select>
        </td>
      <td width="25%" align="right"> <input name="sKey" type="text" class="Input" id="sKey" style="width:100%" value="<%=Trim(Request("sKey"))%>"></td>
      <td width="8%" align="center"> <input name="SearchButton" type="submit" class="button01-out" value="ȷ  ��">
      </td>
    </tr>
  </form>
</table>
<%
Rs.Close
Set Rs=Nothing
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display='';window.scrollTo(window.pageXOffset,2000);}else{HelpTab.style.display='none'}">&nbsp;<img src="Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:none"> 
    <td width="2%">&nbsp;</td>
    <td width="98%" valign="top"> 
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>ҳ������Ƕ�빦�ܽ��ܣ�</td>
        </tr>
        <tr>
          <td>�����˹���ģ����԰�������Ա��վ��ҳ���еĸ�С�����ݽ������߹������³ɾ�̬�ļ��������õķ�Χ�磺ҳ���е�С��桢վ��ͨ�桢��Ȩ���ݼ�����һЩҳ���еı߽����ݿ�<br>
          </td>
        </tr>
      </table>
      <a name="Help"></a></td>
  </tr>
</table>
</body>
</html>
<%
Function ExeSql()
    ExeSql = "Select * From InsertList Where "&sType&" Like '%"&sKey&"%' Order By Id DESC"
End Function
%>