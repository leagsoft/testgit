<%Option Explicit%>
<!--#include file="../../Include/Conn.asp" -->
<!--#include file="upload_5xsoft.asp" -->
<!--#include file="../../Include/Config.asp" -->
<%
Dim Path,Classid,Title
Path=Session("FilePath")
Classid=Request("Classid")
Title=Request("Title")
Dim FilePath,Fso
FilePath=Server.MapPath(DirectoryRoot & Session("FilePath"))
Dim objUpLoad
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
'********************************* Modify By Benny:20040331 ******************************************************
	'If AutoRename Then
		file.SaveAs FilePath &"/"& Year(Now())&Right("00"&Month(Now()),2)&Right("00"&Day(Now()),2)&Right("00"&Hour(Now()),2)&Right("00"&Minute(Now()),2)&Right("00"&Second(Now()),2)&Round(Timer(),0)&"."&File.FileExt
	'Else
	'	file.SaveAs FilePath &"/"& File.FileName
	'End If
'***new****	
	'Dim Filename
	'Filename=Year(Now())&Right("00"&Month(Now()),2)&Right("00"&Day(Now()),2)&Right("00"&Hour(Now()),2)&Right("00"&Minute(Now()),2)&Right("00"&Second(Now()),2)&Round(Timer(),0)&"."&File.FileExt
	'File.SaveAs FilePath &"/"& Filename
'*********************** Modify End ****************************************************************	

Next
Set objUpLoad=Nothing
 
'���ļ������������ݿ⿪ʼ
'Dim Rs,Sql
'Sql="Select Top 1 * From News Order By ID DESC"
'Set Rs=Server.CreateObject ("Adodb.Recordset")
'    Rs.Open Sql,Conn,1,3
'    Rs.AddNew
    '���
'    Rs("Class")=Classid
'    Rs("Title")=Title
'    Rs("Filename")=Filename
'    Rs.Update 
'	Rs.Close 
'set Rs=nothing
'Response.Redirect ("../News_Add.asp?Work=AddReco")
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
if(confirm("<�����ɹ�>\n�ļ��ϴ����,�Ƿ�����ϴ��ļ���(��/��)")){
	window.location="UpFile_Iframe.asp"	
}else{
	top.window.returnValue=true
	top.close()
}
//-->
</SCRIPT>