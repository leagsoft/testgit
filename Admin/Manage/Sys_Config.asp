<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not CBool(SysAdmin.Logined) Then
'    Response.Redirect("Login.asp")
'End If

'If Not SysAdmin.ChangeSysConfig Then
'    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
'    Response.End()
'End If
%>
<html>
<head>
<title>Sys_Config</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link href="Include/ManageStyle.css" rel="stylesheet" type="text/css">
</head>

<body>
<%
Dim Work
    Work=Request("Work")

Select Case Work
    Case "MdyFile"
        Call MdyFile()
    Case "SaveMdy"
        Call SaveMdy()
	 
End Select
%>
<%
Sub MdyFile()
    Dim Fso
    Set Fso= Server.CreateObject(FsoObjectStr)
    Dim sysFile
    Set sysFile=Fso.OpenTextFile(Server.MapPath("Include/Config.asp"))
%>
  <table width="100%" height="489" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
<form name="form1" method="post" action="?">  
    <tr> 
      <td align="center" class="BarTitleBg">ϵͳ�������ã������ã�</td>
    </tr>
    <tr> 
      <td height="410" bgcolor="#FFFFFF"><textarea name="Content" wrap="OFF" class="Input" id="Content" style="width:100%;height:100%"><%=sysFile.ReadAll()%></textarea></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF"><input name="Submit" type="submit" class="button01-out" value="ȷ  ��"> 
        <input name="Submit2" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit3" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"> 
        <input name="Work" type="hidden" id="Work4" value="SaveMdy"  > </td>
    </tr>
</form>    
  </table>
<%
    sysFile.Close()
    Set sysFile=Nothing
    Set Fso=Nothing
End Sub
%>
</body>
</html>
<%
Sub SaveMdy()
    Dim Content
        Content=Request("Content")
    Dim Fso
    Set Fso= Server.CreateObject(FsoObjectStr)
    Dim sysFile
 
    Set sysFile=Fso.OpenTextFile(Server.MapPath("Include/Config.asp"),2,1)

    sysFile.Write Content
    sysFile.Close
    Set Fso=Nothing

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "�޸�ϵͳ����")
    Set LogClass=Nothing

    Response.Write("<script>alert(""<�����ɹ�>\nϵͳ�����������"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End Sub
%>