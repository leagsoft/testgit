<%
	Dim oUpload, NewFileName 
	set oUpload = Server.CreateObject("iNotes.Upload")
	dir = "/Uploads/"	
	oUpload.FilePath=Server.MapPath(dir)
	NewFileUpload=False
	
	If Len(oUpload.Request("File1"))>0 Then

		Randomize 
		strRandom=Int((99999999 * Rnd) + 1)
		tNow=Now()
		strNow=	Year(tNow)&_
			right("0"&Month(tNow),2)&_
			right("0"&day(tNow),2)&_
			right("0"&Hour(tNow),2)&_
			right("0"&Minute(tNow),2)&_
			right("0"&Second(tNow),2)

		'//日期+随机数字作为文件名
		strRealName=strNow&"_"&strRandom
		NewFileName = strRealName&oUpload.ExtName("File1")
		fn=NewFileName
		NewFileUpload=oUpload.SaveFile("File1",NewFileName)
		If NewFileUpload=True Then 
			Response.Write("文件上传成功！")  
		Else 
			Response.Write("文件上传失败！") 
		End If
	Else
		Response.Write ("<script language=javascript>window.close()</script>")
		Response.End 
 	End If
 	set oUpload=Nothing
%> 
<script language=javascript>
function format1(what,opt)
{
	if (opt=="removeFormat")
	{
		what=opt;
		opt=null;
	}

	if (opt==null)
		window.opener.HtmlEditor.document.execCommand(what);
	else
		window.opener.HtmlEditor.document.execCommand(what,"",opt);

	pureText = false;
	window.opener.HtmlEditor.focus();
}

function format(what,opt)
{
	format1(what,opt);
}

	format('InsertImage');
	content=window.opener.HtmlEditor.document.body.innerHTML;

	str="<img src=<%=dir+fn%>>";
	content=content.replace("<IMG>",str);
	window.opener.HtmlEditor.document.body.innerHTML=content;
	window.close();
</script>