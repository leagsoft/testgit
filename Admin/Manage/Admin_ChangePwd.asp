<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/CfsEnCode.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not SysAdmin.Logined Then
'	Response.write("<SCRIPT LANGUAGE=""JavaScript"">alert(""<登录超时>\n请得新登录!"& SoftCopyright_Script &""");window.close()</SCRIPT>")
'	Response.end
'End If

Dim Id,Pwd,Pwd2
Id=Request("Id")
Pwd=Request("Pwd")
Pwd2=Request("Pwd2")
If Pwd<>Pwd2 then
	Response.write("<SCRIPT LANGUAGE=""JavaScript"">alert(""<操作失败>\n两个新密码不一致!"& SoftCopyright_Script &""");</SCRIPT>")
	Response.End
End If
Dim Sql
	If SysAdmin.ChangeAdminList Then
		'如果具有Podm_ChangeAdminList权限
		Sql="Select Top 1 * From Admin Where Id=" & Id
	Else
		'一般管理员只能修改自已的密码
		'If Not SysAdmin.ChagePWD Then
		'	Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
		'	Response.End()
		'End If
		'***************************** Modify By Bennyliu:20040311******************************
		'Sql="Select Top 1 * From Admin Where UCase(Title)='" & UCase(SysAdmin.AdminTitle) &"'"
		Sql="Select Top 1 * From Admin Where Title='" & UCase(SysAdmin.AdminTitle) &"'"
		'****************************************** End Modify *********************************
	End If
Dim Rs
Set Rs=Server.CreateObject("ADODB.RecordSet")	
Rs.Open Sql,Conn,1,3
If Rs.Eof And Rs.Bof Then
	Response.write("<SCRIPT LANGUAGE=""JavaScript"">alert(""<操作失败>\n非法用户"& SoftCopyright_Script &""");</SCRIPT>")
	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
	Response.End	
End If
Rs("Pwd")=CfsEnCode(Pwd)
Rs("UpTime")=Now()
Rs.Update
Rs.Close
Set Rs=Nothing
Conn.Close
Set Conn=Nothing
Response.write("<SCRIPT LANGUAGE=""JavaScript"">alert(""<操作成功>\n密码修改成功"& SoftCopyright_Script &""");top.close();</SCRIPT>")
%>