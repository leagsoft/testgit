<!--#include file="../Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="../Include/Config.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not SysAdmin.Logined Then
'    Response.Write("<script>alert(""<����ʧ��>\n������ʱ����δ��¼"& SoftCopyright_Script &""");top.window.close();</script>")
'    Response.End()
'End If

'If Not SysAdmin.ManageFiles Then
'    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");top.window.close();</script>")
'    Response.End()
'End If

Session("FilePath")=Trim(Request("Path"))
%>
<html>
<head>
<title>�ļ��ϴ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="buttonface" leftmargin="0" topmargin="0">
<iframe width="100%" height="100%" frameborder="0" src="UpFile_Iframe.asp"></iframe>
</body>
</html>