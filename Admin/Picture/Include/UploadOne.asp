<%Response.End %>
<!--#Include File="Inc/Function.asp"-->
<%
On Error Resume Next
'�����ϴ��ļ���
'*******************Modify by BennyLiu:20030410*********************
cFolder = Server.MapPath("/Upload")
'cFolder="d:\www\chinapackaging-union.com\EN\Upload"
'cFolder="d:\WorkRoom\www\chinapackaging-union.com\EN\Upload"
'cFolder = Server.MapPath("/EN/Upload")
'**************************Modify End******************************

action=Request("action")
Randomize 
sf=Int((99999999 * Rnd) + 1)

If action="upload" Then
  '�ϴ��ļ�
  Set oUpload = Server.CreateObject("iNotes.Upload")
  oUpload.FilePath=cFolder
  newfileupload=false
  If UCase(oUpload.ExtName("file"))<>".JPG" Then
     cMsg = "��Ч�ļ���ϵͳ��֧��JPG��ʽ���ļ��ϴ���������ѡ��<br>"
  ElseIf oUpload.FileSize("file")>204800 Then
     cMsg = "���ϴ����ļ�����200K����ѹ�������ϴ���<br>"
  Else
     filename=oUpload.FileName("file")
     newfileupload=oUpload.SaveFile("file","gallery"&sf&UCase(oUpload.ExtName("file")))
     cMsg = "�ļ�("&filename&")�ϴ��ɹ���<br>"
     Session("cApply")="gallery"&sf&UCase(oUpload.ExtName("file"))

     '����Ԥ��ͼƬ
     Set objPictureProcessor = Server.CreateObject("COMobjects.NET.PictureProcessor")
     objPictureProcessor.LoadFromFile cFolder&"\"&Session("cApply")
     objPictureProcessor.OptimizationOn = True
     objPictureProcessor.Quality = 90
     intNewWidth = 150
     intNewHeight=CInt(objPictureProcessor.Height*150/objPictureProcessor.Width)
     If intNewHeight > 150 Then intNewHeight = 150
     objPictureProcessor.Resize intNewWidth,intNewHeight
     objPictureProcessor.SaveToFileAsJpeg cFolder&"\Small\"&Session("cApply")
     Set objPictureProcessor = Nothing
     
     If newfileupload=false Then
        cType = "ʧ��"
        cMsg="�ļ�("+filename+")�ϴ�ʧ�ܣ������룺INC_UPLOADONE_001<br>"'//+oUpload.Error
        Call PutEvent(cType,Session("cUserId"),cMsg,"","","")
     End If
  End If
  Set oUpload = nothing
ElseIf action="delete" Then
  Dim fso, MyFile
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set MyFile = fso.GetFile(Trim(cFolder&"\"&Session("cApply")))
  MyFile.Delete
  Set MyFile2 = fso.GetFile(Trim(cFolder&"\Small\"&Session("cApply")))
  MyFile2.Delete   
  MyFile.Close
  MyFile2.Close
  fso.Close 
  cMsg = "�ļ�("&Session("cApply")&")ɾ���ɹ���<br>"
  Session("cApply")=""
End If
If Session("cApply") = "" Then
   cAction = "upload"
Else
   cAction = "delete"
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link REL="stylesheet" href="/Css/Font.css" TYPE="text/css">
<title>ͼƬ�ϴ���</title>
<script language="javascript">
function finish()
{
  opener.document.all.nPic.value=document.all.cPic.value;
  window.close();
}
</script> 
</head>
<body bgcolor="#0D66AE" leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0 onload="javascript:window.focus()">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<form method=post ENCTYPE="multipart/form-data" action="UploadOne.asp?action=<%=cAction%>">
<input type=hidden value="<%=Session("cApply")%>" name="cPic">
<tr height="50">
<td align="center">
<input type="hidden" name="Copyright" value="�й�[���ʱ���]����,http://China-Notes.com">
<input type="file" size=30 name="file"<%If Session("cApply")<>"" Then Response.Write " disabled"%>>&nbsp;
<input type=submit value="�ϴ�"<%If Session("cApply")<>"" Then Response.Write " disabled"%>>
<input type=submit value="ɾ��"<%If Session("cApply")="" Then Response.Write " disabled"%>>
<input type=button onclick="javascript:finish()" value="���">
</td>
</tr>
<tr>
<td bgcolor="#FDB85B" height="300" align="center" class=bigfont><font color=red><%=cMsg%></font>
<%
If Session("cApply")<>"" Then 
'******************************Modify by BennyLiu:20030410*********************
  'Response.Write "<img src=/Upload/Small/"&Session("cApply")&" border=1>"
   Response.Write "<img src=/EN/Upload/Small/"&Session("cApply")&" border=1>"
'*********************************MOdify End *******************************
Else
   Response.Write "ͼƬԤ����"
End If
%>
</td>
</tr>
</form>
</table>
