<%Option Explicit%>
<!--#include file="upload_5xsoft.asp" -->
<!--#include file="../../Include/Config.asp" -->
<%
Dim FilePath,Fso
FilePath=Server.MapPath(DirectoryRoot & "")

Dim objUpLoad
Dim cFileName
Set objUpLoad = new upload_5xsoft
Dim formName,File,AutoRename
	AutoRename=objUpLoad.Form("AutoRename")="1"
For Each formName In objUpLoad.objFile
    Set File=objUpLoad.objFile(formName)
    '�Ƿ���������ļ�
    Set Fso = Server.CreateObject(FsoObjectStr)

    If AutoRename=False And Fso.FileExists(FilePath &"/"& File.FileName) Then
        Set Fso=Nothing
        Response.Write("<script>alert(""<����ʧ��>\n����ͬ���ļ�"");"&_
                        "window.location=""UpFile_Iframe.asp""</script>")
        Response.End
	End If

	If InStr(FileSystem_EnableFileExt,"|"&UCase(File.FileExt)&"|")=0 Then
		Response.Write("<script>alert(""<����ʧ��>\n�ļ����Ͳ�������"");"&_
						"window.location=""UpFile_Iframe.asp""</script>")
		Response.End
	End If
	
	'If AutoRename Then
		file.SaveAs FilePath &"/"& Year(Now())&Right("00"&Month(Now()),2)&Right("00"&Day(Now()),2)&Right("00"&Hour(Now()),2)&Right("00"&Minute(Now()),2)&Right("00"&Second(Now()),2)&Round(Timer(),0)&"."&File.FileExt
		cFileName=Year(Now())&Right("00"&Month(Now()),2)&Right("00"&Day(Now()),2)&Right("00"&Hour(Now()),2)&Right("00"&Minute(Now()),2)&Right("00"&Second(Now()),2)&Round(Timer(),0)&"."&File.FileExt
		
	'Else
	'	file.SaveAs FilePath &"/"& File.FileName
	'	cFileName=File.FileName	
	'End If

Next
Set objUpLoad=Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--	
	alert("�ļ��ϴ��ɹ���");
	top.window.returnValue="<%=cFileName%>"
	top.close ();
//-->
</SCRIPT>