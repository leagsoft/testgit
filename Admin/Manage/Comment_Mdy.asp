<!--#include file="../Comment/conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not CBool(SysAdmin.Logined) Then
'    Response.Redirect("Login.asp")
'End If

'If Not SysAdmin.ChangeCommentList Then
'    Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
'    Response.End()
'End If
%>
<html>
<head>
<title>Comment_Mdy.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
</head>
<body>
<%
Select Case Request("Work")
    Case "SaveMdy"
        Call SaveMdy()
    Case "DelReco"
        Call DelReco()
    Case "AddReco"
        Call AddReco()
    Case Else
End Select
%>
</body>
</html>
<%
Sub DelReco()
    Dim Sql
    Dim ItemChkBox
    ItemChkBox=Request("ItemChkBox")
    If ItemChkBox="" Then
        Response.Write("<script>alert(""<操作失败>\n没有选择删除项目"");window.history.back()</script>")
        Response.end
    End If
    Sql="Delete From VisitorComment Where Id In (" & ItemChkBox & ")"
    Conn.ExeCute(Sql)
    Response.Redirect("Comment_List.asp")
End Sub
%>