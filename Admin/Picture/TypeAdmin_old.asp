<!--#include File="../../Include/Conn.asp"-->
<%
Work=Request("Work")
'���尴ť
cAdd    = "����"
cReset  = "����"
cFinish = "���"
cDel    = "ɾ��"
cSave   = "�޸�"

Action=Trim(Request("cAction"))
'���ӷ���
IF Action="����" then
	cValue=Trim(Request("cValue"))
	cIspicture=Request("IsPicture")
	if cIspicture="" then
		cIsPicture=0
	end if
	Sql="insert into SYSDIC (Type,IsPicture) values ('"&cValue&"','"&cIsPicture&"')"
	Set Rs1=Server.CreateObject ("Adodb.Recordset")
	Rs1.Open Sql,Conn,1,3 
	Set Rs1=nothing
	Response.Write "<script language='javascript'>alert('�������ɹ�');</script>"
End IF

IF Work="01" then
	cDicId=Trim(Request("cDicId"))
	Sql="select * from SYSDIC where DICID="&cDicId
	Set Rs2=Server.CreateObject ("Adodb.Recordset")
	Rs2.Open Sql,conn,1,3
	cValue=Rs2("Type")
	Rs2.Close 
	set Rs2=nothing
End IF

'�޸ķ���
IF Action="�޸�" then
	cDicId=Trim(Request("cDicId"))
	IF cDicId="" then
		Response.Redirect ("TypeAdmin.asp")
	Else
	cValue=Trim(Request("cValue"))
	cIsPicture=Trim(Request("Ispicture"))
	If cIspicture="" then
		cIsPicture=0
	End If
	Sql="Update SYSDIC set Type='"&cValue&"',IsPicture='"&cIsPicture&"' where DicId="&cDicId
	'Response.Write sql
	'Response.End
	Set Rs3=Server.CreateObject ("Adodb.Recordset")
	Rs3.Open Sql,conn,1,3
	'Rs3.Close 
	set Rs3=nothing
	Response.Write "<script language='javascript'>alert('�޸����ɹ�');</script>"
	End IF
	
End IF

'ɾ��
IF Action="ɾ��" Then
	cDicId=Trim(Request("cDicId"))
	IF cDicID="" then
		Response.Redirect ("TypeAdmin.asp")
	Else
		cValue=Trim(Request("cValue"))
		Sql="Delete From SYSDIC where DICID="&cDicId
		Set Rs4=server.CreateObject ("Adodb.Recordset")
		Rs4.Open Sql,conn,1,3
		Set Rs4=nothing
	End IF
End IF
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������</title>
</head>
<link rel="stylesheet" href="GreatSoft.css" type="text/css">
<body marginheight="5" marginwidth="0" topmargin="5" leftmargin="10" rightmargin="0" background="images/lay2_main_bg.gif">

<form METHOD="POST" name="TypeAdmin" action="TypeAdmin.asp">
<!--<input type="hidden" name="cMode" value="<%= cMode%>">-->
<input type="hidden" name="cDicId" value="<%= cDicId%>">
<!--<input type="hidden" name="cType" value="<%= cType%>">-->
<div align="center">
<table width="95%" border="1" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFDFDF" bordercolorlight="#333333" bordercolordark="#FFFFFF">
<!--<tr>
 <td bgcolor="#FDD590" height="30" colspan="2">&nbsp;<%=cDisp%>&nbsp;<font color=red><%=cMsg%></font></td>
</tr>-->
<tr>
<td bgcolor="#ccffcc" height="20" width="20%" align="right">�������ƣ�</td>
<td bgcolor="#ffffff" height="20"><input type="text" name="cValue" size="40" value=<%=cValue%>></td>
</tr>
<tr>
<td bgcolor="#ccffcc" height="20" width="20%" align="right">�����Ƿ����ͼƬ��</td>
<td bgcolor="#ffffff" height="20"><input type="checkbox" name="Ispicture" value="1" checked></td>
</tr>
</table>
</div>	
<div align="center"><br>
<input type="submit" name="cAction" value="<%= cSave%>">&nbsp;&nbsp;<input type="submit" name="cAction" value="<%= cAdd%>">&nbsp;&nbsp;<input type="submit" name="cAction" value="<%= cDel%>">&nbsp;&nbsp;
<!--<input type="submit" name="cAction" value="<%= cSave%>">&nbsp;&nbsp;<input type="submit" name="cAction" value="<%= cReset%>">&nbsp;&nbsp;-->
<input type="button" name="Submit" value="���" onclick="javascript:window.close()">
<br><br>
</div>
<div align="center">
<table width="95%" border="1" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFDFDF" bordercolorlight="#333333" bordercolordark="#FFFFFF">
<tr>
 <td bgcolor="#FDD590" height="30" colspan="2">&nbsp;���еķ���</td>
</tr>
<tr>
<td BGCOLOR="#339900" height="20" align="center" width="80%"><font color="#ffffff">��������</font></td>
<td BGCOLOR="#339900" height="20" align="center" width="20%"><font color="#ffffff">�����Ƿ������ͼƬ</font></td>
</tr>
<%
'��ѯ�Ѷ�������
cSql = "select * from SYSDIC where DELETED=0 order by DICID asc"
Set Rs=Server.CreateObject ("Adodb.Recordset")
Rs.Open cSql,Conn,1,3
Do
  If Rs.Eof Then Exit Do
%>
<tr>
<td height="22" width="80%" align="center"><a href="TypeAdmin.asp?Work=01&cDicId=<%=Rs("DICID")%>"><%= Rs("Type")%></a></td>
<td height="22" width="20%" align="center"><input type="checkbox" name="checkbox" <%if Rs("Ispicture")="1" then Response.Write "checked" end if%>></td>
</tr>
<%
  Rs.MoveNext
Loop
'�رն���
Rs.Close
Set Rs = Nothing
Conn.Close
Set Conn = Nothing
'�ر����� 
%>
</table>
</div>
</form>
</body>
</html>

