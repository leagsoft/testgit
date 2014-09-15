<!--#include file="../../Include/Conn.asp" -->
<!--#include file="upload_5xsoft.asp" -->
<!--#include file="../../Include/Config.asp" -->
<%
on error resume next
  
Dim Path,Classid,Title,EdiTor,UserBM
Path=Session("FilePath")
Classid=Request("Classid")
Title=Request("Title")
Editor=Session("YHDL")	'编辑者
UserBM=Session("YHBM")	'用户部门
Author=Trim(Reqeust("Author"))		'*****Add By BennyLiu:20040618 ********
Browser=Session("Browser")			'Add By BennyLiu:20040625 增加浏览者
DocumentType=Session("DocumentType")	'Add By BennyLiu:20040706 增加文件类型
'Response.Write Browser
'Response.End 
Dim FilePath,Fso
FilePath=Server.MapPath(DirectoryRoot & Session("FilePath"))
Dim objUpLoad
Set objUpLoad = new upload_5xsoft
Dim formName,File,AutoRename
	AutoRename=objUpLoad.Form("AutoRename")="1"
For Each formName In objUpLoad.objFile
    Set File=objUpLoad.objFile(formName)
    '是否存在重名文件
    Set Fso = Server.CreateObject(FsoObjectStr)

    If AutoRename=False And Fso.FileExists(FilePath &"/"& File.FileName) Then
        Set Fso=Nothing
		%>
		 
		 
			<script language="vbscript">
			intResponse=msgbox("存在同名文件,是否继续上传文件",1,"提示")
			 
			if intResponse=1 then
			
			else
					window.history.back()
					
			end if
			
			</script>
			
	 
		<%
 
	End If
 
	If InStr(FileSystem_EnableFileExt,"|"&UCase(File.FileExt)&"|")=0 Then
		Response.Write("<script>alert(""<操作失败>\n文件类型不被允许"");"&_
						"window.location=""UpFile_Iframe.asp""</script>")
		Response.End
	End If
'********************************* Modify By Benny:20040331 ******************************************************
	'If AutoRename Then
	'	file.SaveAs FilePath &"/"& Year(Now())&Right("00"&Month(Now()),2)&Right("00"&Day(Now()),2)&Right("00"&Hour(Now()),2)&Right("00"&Minute(Now()),2)&Right("00"&Second(Now()),2)&Round(Timer(),0)&"."&File.FileExt
	'Else
	'	file.SaveAs FilePath &"/"& File.FileName
	'End If
'***new****	
	'*Dim Filename
	'*Filename=Year(Now())&Right("00"&Month(Now()),2)&Right("00"&Day(Now()),2)&Right("00"&Hour(Now()),2)&Right("00"&Minute(Now()),2)&Right("00"&Second(Now()),2)&Round(Timer(),0)&"."&File.FileExt
	'*File.SaveAs FilePath &"/"& Filename
	Dim Filename
		If AutoRename Then
			Filename=Year(Now())&Right("00"&Month(Now()),2)&Right("00"&Day(Now()),2)&Right("00"&Hour(Now()),2)&Right("00"&Minute(Now()),2)&Right("00"&Second(Now()),2)&Round(Timer(),0)&"."&File.FileExt
		Else
			Filename=File.FileName
		End If
		 
		File.SaveAs FilePath &"/"&Filename
'*********************** Modify End ****************************************************************	

Next
Set objUpLoad=Nothing

'把文件名保存入数据库开始
Dim Rs,Sql
Sql="Select Top 1 * From News Order By ID DESC"
Set Rs=Server.CreateObject ("Adodb.Recordset")
    Rs.Open Sql,Conn,1,3
    Rs.AddNew
    '入库
    Rs("Class")=Classid
    Rs("Title")=Title
    Rs("Editor")=Editor
    Rs("Author")=Author
    Rs("Filename")=Filename
    Rs("UserBM")=UserBM
    if Browser<>"" then
		Rs("Browser")=Browser
    end if
    if DocumentType<>"" then
		Rs("IsDocument")=DocumentType
    end if
    Rs.Update 
	Rs.Close 
set Rs=nothing
Session("cBrowser")=Browser
 Response.Cookies("ZGW_NewsSys3")("CurrentClassIdUsed")=ClassId
%>
<script>
window.location.href="../News_Add.asp?Work=AddReco"
</script>
