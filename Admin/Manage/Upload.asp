<!--#include file="inc/upfile_class.asp"-->
<%
Server.ScriptTimeOut = 1800
Dim sAllowExt, nAllowSize, sUploadDir
dim themax
sAllowExt = UCase("TXT|HTML|HTM|TXT|GIF|JPG|JPEG|BMP|PNG|DOC|XLS|TIF|RAR|ZIP|EXE|MHT|PPT|MP3|CHM|SWF|RM|FLV|")
nAllowSize=999999
sUploadDir="../UpLoadFiles"
'**Call InitUpload()		' 初始化上传变量

Dim sAction
sAction = UCase(Trim(Request.QueryString("action")))

Call ShowForm()			' 显示上传表单
If sAction = "SAVE" Then
	Call DoSave()		' 存文件
End If



Sub ShowForm() 
%>
<HTML>
<HEAD>
<TITLE>文件上传</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<Link rel="stylesheet" type="text/css" href="pop.css">
<style type="text/css">
body, a, table, div, span, td, th, input, select{font:9pt;font-family: "宋体", Verdana, Arial, Helvetica, sans-serif;}
body {padding:0px;margin:0px}
</style>

<script language="JavaScript" src="wbtextbox/dialog.js"></script>

</head>
<body bgcolor=menu>

<form action="?action=save" method=post name=myform enctype="multipart/form-data">
<input type=file name=uploadfile size=1 style="width:100%">
<%
'**response.Write "可上传的文件："&sAllowExt&"<br>单个文件大小："&nAllowSize&"KB　可用空间："&themax&"KB"
'response.Write "可上传的文件："&sAllowExt&"<br>单个文件大小："&nAllowSize&"KB"
%>
</form>

<script language=javascript>

var sAllowExt = "<%=sAllowExt%>";
// 检测上传表单
function CheckUploadForm() {
	if (!IsExt(document.myform.uploadfile.value,sAllowExt)){
		parent.UploadError("提示：\n\n请选择一个有效的文件，\n支持的格式有（"+sAllowExt+"）！");
		return false;
	}
	return true
}

// 提交事件加入检测表单
var oForm = document.myform ;
oForm.attachEvent("onsubmit", CheckUploadForm) ;
if (! oForm.submitUpload) oForm.submitUpload = new Array() ;
oForm.submitUpload[oForm.submitUpload.length] = CheckUploadForm ;
if (! oForm.originalSubmit) {
	oForm.originalSubmit = oForm.submit ;
	oForm.submit = function() {
		if (this.submitUpload) {
			for (var i = 0 ; i < this.submitUpload.length ; i++) {
				this.submitUpload[i]() ;
			}
		}
		this.originalSubmit() ;
	}
}

// 上传表单已装入完成
try {
	parent.UploadLoaded();
}
catch(e){
}

</script>

</body>
</html>
<% 
End Sub 

' 保存操作
Sub DoSave()
	Dim oUpload, oFile, sFileExt, sFileName
	dim osize,username,rs
	' 建立上传对象
	Set oUpload = New upfile_class
	' 取得上传数据,限制最大上传
	
	
	'Call Checksize(osize)
	
	oUpload.GetData(nAllowSize*1024)
	If oUpload.Err > 0 Then
		select Case oUpload.Err
		Case 1
			Call OutScript("parent.UploadError('请选择有效的上传文件！')")
		Case 2
			Call OutScript("parent.UploadError('你上传的文件总大小超出了最大限制（" & nAllowSize & "KB）！')")
		End Select
		Response.End
	End If
	
	Set oFile = oUpload.File("uploadfile")
	sFileExt = UCase(oFile.FileExt)
	osize = oFile.Filesize
	Call CheckValidExt(sFileExt)


	Dim sRnd
	Randomize
	sRnd = Int(900 * Rnd) + 100
	sFileName = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & sRnd & "." & sFileExt	
	'response.write sfilename
	'Response.End 
	oFile.SaveToFile Server.Mappath(sUploadDir & "/"& sFileName)
	
	Set oFile = Nothing
	Set oUpload = Nothing
	
	'if Request.Cookies("oblog")("UserName")<>"" then
	'	username=trim(Request.Cookies("oblog")("UserName"))
	'	set rs=conn.execute("select upfiles from [user] where username='"&username&"'")
	'	if not rs.eof then
	'		if rs(0)<>"" then
	'			conn.execute("update [user] set upfiles='"&rs(0)&"|"&sFileName&"' where username='"&username&"'")	
	'		else
	'			conn.execute("update [user] set upfiles='"&sFileName&"' where username='"&username&"'")	
	'		end if
	'		conn.execute("update [user] set upfiles_size=upfiles_size+"&osize&" where username='"&username&"'")
	'	end if
	'end if
	'set rs=nothing
	'call closeconn()
	
	Call OutScript("parent.UploadSaved('" & sUploadDir & "/"& sFileName & "')")	

End Sub

' 输出客户端脚本
Sub OutScript(str)

	Response.Write "<script language=javascript>" & str & ";history.back()</script>"
End Sub

' 检测扩展名的有效性
Sub CheckValidExt(sExt)
	Dim b, i, aExt
	b = False
	aExt = Split(sAllowExt, "|")
	For i = 0 To UBound(aExt)
		If UCase(aExt(i)) = UCase(sExt) Then
			b = True
			Exit For
		End If
	Next
	If b = False Then
		OutScript("parent.UploadError('提示：\n\n请选择一个有效的文件，\n支持的格式有（"+sAllowExt+"）！')")
		Response.End
	End If
End Sub

Sub Checksize(osize)
	If osize>nAllowSize*1024 Then
		call OutScript("parent.UploadError('你上传的文件总大小超出了最大限制（" & nAllowSize & "KB）！')")
		Response.End
	End If
End Sub


' 初始化上传限制数据

%>