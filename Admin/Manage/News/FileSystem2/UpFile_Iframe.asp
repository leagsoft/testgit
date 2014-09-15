<!--#include file="../../Include/Config.asp" -->
<%
Dim Title,Classid
Classid=Request("Classid")
Title=Request("Title")
%>
<html>
<head>
<title>文件上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="buttonface" leftmargin="0" topmargin="0" scroll="no">
<table width="1019" border="0" cellspacing="0" cellpadding="5" height="169">
  <tr> 
    <td height="196" valign="middle" bgcolor="#FFFFFF" width="153" align="center"> 
      <h3> <img src="images/UpFileLogo.gif" width="130" height="131"></h3>
    </td>
    <td height="196" width="846"> 
      <table width="464" border="0" cellspacing="1" cellpadding="2">
        <form name="form1" method="post" enctype="multipart/form-data" action="UpFile_SaveFile.asp?Title=<%=Title%>&Classid=<%=Classid%>" onSubmit="return ChkForm(this)">
          <input name="Path" type="hidden" id="Path2" value="<%=Request("Path")%>">
          <tr> 
            <td width="16%" align="right"><font size="2"><b>文件上传</b></font>:</td>
            <td width="84%"> 
              <input name="File1" type="File" size="20">
            </td>
          </tr>
          <tr> 
            <td width="16%">&nbsp;</td>
            <td width="84%"> 
              <input name="Submit" type="Submit" value="开始上传">
              <input name="Submit2" type="button" value="取消" onClick="closeWin()">
              <!--<input name="Path" type="hidden" id="Path2" value="<%=Request("Path")%>">-->
              <script language="JavaScript" type="text/JavaScript">
function closeWin()
{
	top.returnValue=false;
	top.window.close();
}		
function ChkForm(Obj)
{
	if(Obj.File1.value==""){
		alert("请选择文件");
		return false;
	}
	Obj.Submit.disabled=true;
	return true;
}
</script>
            </td>
          </tr>
          <!--<tr> 
            <td width="16%">&nbsp;</td>
            <td width="84%"> <label for="AutoRename">
              <input type="checkbox" name="AutoRename" value="1" id="AutoRename" checked>
              自动重命名文件</label></td>
          </tr>-->
        </form>
      </table>
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="2"> 
            <hr size="1" width="100%">
          </td>
        </tr>
        <tr> 
          <td height="54" valign="middle"> 
            <li>文件上传速度较慢，请耐心等侍</li>
            <li>自动命名的格式为：2003100522040479444.gif</li>
			<li>允许上传文件类型：<%=FileSystem_EnableFileExt%></li>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  
</body>
</html>
