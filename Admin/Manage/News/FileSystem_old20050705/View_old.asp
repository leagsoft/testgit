<%Option Explicit%>
<!--#include file="FunLib.asp" -->
<!--#include file="../../Include/Config.asp" -->
<!--#include file="../../Include/Tkl_SYSProedomClass.asp" -->
<%
dim classid,title
Classid=Request("radioBoxItem")
Title=Request("Title")
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not SysAdmin.Logined Then
 '   Response.Redirect("../../Login.asp")
'End If
%>
<%
    Dim tempCurrentPath,CurrentPath
        tempCurrentPath=FilterPath(Trim(Request("CurrentPath")))    
    If tempCurrentPath="" Then
        CurrentPath=DirectoryRoot
    Else
        CurrentPath=DirectoryRoot & tempCurrentPath
    End If
	

    Dim Fso
    Set Fso = Server.CreateObject(FsoObjectStr)
    Dim Fol,Fols    'Folder,Folders
    Dim Fle,Fles    'File,Files
    If Fso.FolderExists(Server.MapPath(CurrentPath)) Then    
        '��������ļ�ϵͳ��Ŀ¼DirectoryRoot
        Set Fol=Fso.GetFolder(Server.MapPath(CurrentPath))        
    Else
        '�������ļ�ϵͳ��Ŀ¼δ�ҵ�,�򴴽���Ŀ¼
        If CurrentPath=DirectoryRoot Then
            Fso.CreateFolder(Server.MapPath(CurrentPath))
            Set Fol=Fso.GetFolder(Server.MapPath(CurrentPath))
        End If
    End If
    
    Select Case Request("Work")
        Case "DelItem"
            DelItem()
            Response.Redirect("?CurrentPath=" & tempCurrentPath)
        Case "MoveItem"
            MoveItem()
            Response.Redirect("?CurrentPath=" & tempCurrentPath)
        Case "CreateFolder"
            CreateFolder()
            Response.Redirect("?CurrentPath=" & tempCurrentPath)
        Case Else
    End Select

    Dim ItemCount
    ItemCount=0
%>
<html>
<head>
<title>����������ļ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Style.css" rel="stylesheet" type="text/css">
</head>

<body>
<a name="Top"></a> 
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr> 
    <td width="15%"> <img src="images/m_Home.gif" width="20" height="20" style="cursor:hand" onClick="location='?'" title="�ص���Ŀ¼"><img src="images/m_Front.gif" width="20" height="20" style="cursor:hand" onClick="history.back()" title="����"><img src="images/m_Back.gif" width="20" height="20" style="cursor:hand" onClick="history.forward()" title="ǰ��"><img src="images/m_STOP.gif" width="20" height="20" style="cursor:hand" onClick="window.stop" Title="ֹͣ"><img src="images/m_Ref.gif" width="20" height="20" style="cursor:hand" onClick="location.reload()" title="ˢ��"><!--<img src="images/m_MakeFolder.gif" width="20" height="20" style="cursor:hand" onClick="CreateFolder()" title="����Ŀ¼">--><img src="images/m_NewFile.gif" width="20" height="20"  style="cursor:hand" onClick="UploadFile()" title="�ϴ��ļ�"></td>
    <td>
      <input name="Path" type="text" class="Input" id="Path" value="<%=Replace(CurrentPath,DirectoryRoot,"��Ŀ¼")%>" size="80"  readonly="true" style="width:100%">
    </td>
  </tr>
</table>
<form name="form1" method="post" action="">
  <table border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" style="table-layout:fixed; word-break :break-all;width:100%">
    <tr bgcolor="#666699"> 
      <td width="4%" align="center">&nbsp;</td>
      <td width="60%"><font color="#FFFFFF">&nbsp;&nbsp;����</font></td>
      <td width="14%" align="center"><font color="#FFFFFF">&nbsp;&nbsp;��С</font></td>
      <td width="13%" align="center" title="Created Time"><font color="#FFFFFF"> 
        ����ʱ��</font></td>
      <td width="9%" align="center"><font color="#FFFFFF">����</font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="center"><img src="images/Folder_Root.gif" width="16" height="16" Title="�ļ���"></td>
      <td><a <%If tempCurrentPath<>"" Then%>href="?CurrentPath=" <%Else%>href="#" disabled="true"<%End If%>>�ص���Ŀ¼</a></td>
      <td align="center"><font color="#999999">None</font></td>
      <td align="center"><font color="#999999">None</font></td>
      <td align="center"><font color="#999999">None</font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="center"><img src="images/Folder_Parent.gif" width="16" height="16" Title="�ļ���"></td>
      <td><a <%If tempCurrentPath<>"" Then%>href="?CurrentPath=<%=GetParent(tempCurrentPath)%>" <%Else%>href="#" disabled="true"<%End If%>>�ص���Ŀ¼</a></td>
      <td align="center"><font color="#999999">None</font></td>
      <td align="center"><font color="#999999">None</font></td>
      <td align="center"><font color="#999999">None</font></td>
    </tr>
    <%
    Call ShowFolderList()
    Sub ShowFolderList()
        Dim Item
        Set Fols=Fol.SubFolders
        For Each Item In Fols 
            ItemCount=ItemCount+1
%>
    <tr bgcolor="#FFFFFF"> 
      <td align="center"><img src="images/Folder.gif" width="16" height="16" Title="�ļ���"></td>
      <td><a href="?CurrentPath=<%=tempCurrentPath&"/"&Item.Name%>"><%=Item.Name%></a></td>
      <td align="center"><%=FormatNumber(Item.Size/1024,2)&" KB"%></td>
      <td align="center"><span Title="<%=Item.DateCreated%>"><%=FormatDateTime(Item.DateCreated,2)%></span></td>
      <td align="center"> 
        <input name="Item" type="checkbox" id="Item" value="<%=Item.Name & "1"%>">
      </td>
    </tr>
    <%
        Next
    End Sub
%>
    <%
    Call ShowFileList()
    Sub ShowFileList()
        Dim Item
        Set Fles=Fol.Files
        For Each Item In Fles 
            ItemCount=ItemCount+1
%>
    <tr bgcolor="#FFFFFF"> 
      <td align="center"><img src="images/<%=FileIco(Item.Name)%>" width="16" height="16"></td>
      <td bgcolor="#FFFFFF"><a href="<%=DirectoryRoot&tempCurrentPath&"/"&Item.Name%>" target="_balnk" Title="<%=Item.Type%>" Id="File<%=ItemCount%>"><%=Item.Name%></a> 
        <!--<a href="#" onclick="prompt('[<%=Item.Name%>]�ļ���URL',File<%=ItemCount%>.href)"><font color="#0000FF">ȡ��URL</font></a>]--></td>
      <td align="center"><%=FormatNumber(Item.Size/1024,2)&" KB"%></td>
      <td align="center" colspan="2"><span Title="<%=Item.DateCreated%>"><%=FormatDateTime(Item.DateCreated,2)%></span></td>
      <!--<td align="center"> 
        <input name="Item" type="checkbox" id="Item" value="<%=Item.Name & "0"%>">
      </td>-->
    </tr>
    <%
        Next
    End Sub
%>
    <tr align="right" bgcolor="#f6f6f6"> 
      <td colspan="5"> Total:<%=ItemCount%>&nbsp;|&nbsp; <a href="#Top">Top</a></td>
    </tr>
  </table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
    <tr>
      <td align="right">
        <input type="submit" name="Submit" value="Submit" style="display:none">
        <input name="Work" type="hidden" id="Work">
        <input name="Parameter" type="hidden" id="Parameter">
        <input name="CurrentPath" type="hidden" id="CurrentPath" value="<%=tempCurrentPath%>">
        <!--<img src="images/DelFile.gif" width="36" height="37" onClick="DelItem()" title="ɾ����ѡ��Ŀ" style="cursor:hand"><img src="images/MoveFile.gif" width="36" height="37" onClick="MoveItem()" title="�ƶ���ѡ��Ŀ" style="cursor:hand">--> 
      </td>
    </tr>
  </table>
</form>
<script language="JavaScript" type="text/JavaScript">
function UploadFile()
{
    //var Result=showModalDialog("UpFile.asp?Title=<%=Title%>&Classid=<%=Classid%>&Path=<%=tempCurrentPath%>","","dialogWidth:500px;dialogHeight:230px;center:yes;scroll:no;");
    //if(Result){
        //window.location.reload();
   // }
   window.location .href("UpFile.asp?Title=<%=Title%>&Classid=<%=Classid%>&Path=<%=tempCurrentPath%>")
}
function DelItem()
{
    if(confirm("��ȷ��Ҫ����ѡ�е��ļ�/�ļ���ɾ����\n(ɾ�����޷���ԭ)"))
    {    
        form1.Work.value='DelItem';
        form1.Submit.click()
    }
}
function MoveItem()
{
    var Result=prompt("������Ŀ��Ŀ¼λ��(�������)��","")
    if(Result)
    {    
        form1.Work.value="MoveItem"
        form1.Parameter.value=Result
        form1.Submit.click()
    }
}
function CreateFolder()
{
    var str;
    str=prompt("<����Ŀ¼>\n������[Ŀ¼��]:","Folder"+Math.floor(Math.random()*1000));
    if(str!=null)
    {
        if(str=="")
        {
            alert("<����ʧ��>\nδ��д[Ŀ¼��]");
            return false;
        }else{
            window.location='?Title=<%=Title%>&Classid=<%=Classid%>&CurrentPath=<%=tempCurrentPath%>&Work=CreateFolder&Title='+str;
        }
    }
}
</script>
<br>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display='';window.scrollTo(window.pageXOffset,2000);}else{HelpTab.style.display='none'}">&nbsp;<img src="../../Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:"> 
    <td width="2%">&nbsp;</td>
    <td width="98%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td>��ɾ��</td>
        </tr>
        <tr> 
          <td>������ɾ�����ļ�/Ŀ¼������Ŀ¼�µ�������Ŀ¼���ļ��������ܻ�ԭ</td>
        </tr>
        <tr> 
          <td>���ƶ�</td>
        </tr>
        <tr> 
          <td>Ŀ��Ŀ¼���Ľ�β���ܰ�����/����</td>
        </tr>
        <tr> 
          <td>��������ֻ���룢/����ʾ�ƶ�����Ŀ¼</td>
        </tr>
        <tr> 
          <td>����������/aaa/bbb����ʾ�ƶ�������Ŀ¼��ʼ�����aaa/bbbĿ¼����</td>
        </tr>
        <tr>
          <td>����������aaa/bbb����ʾ�ƶ����ӵ�ǰ����Ŀ¼��ʼ�����aaa/bbbĿ¼����</td>
        </tr>
      </table>
      <a name="Help"></a></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
Set Fol=Nothing
Set Fso=Nothing

'########ȡ���ļ���չ������Ӧ��ͼ��
Function FileIco(f_name)
    Dim ex,ico
    ex=LCase(GetEx(f_name))
    Select Case ex
        Case ".doc"
            ico="f_Doc.gif"
        Case ".txt"
            ico="f_txt.gif"
        Case ".mp3"
            ico="f_mp3.gif"
        Case ".gif"
            ico="f_pic.gif"
        Case ".bmp"
            ico="f_pic.gif"
        Case ".jpg"
            ico="f_pic.gif"
        Case ".ico"
            ico="f_pic.gif"
        Case ".rar"
            ico="f_rar.gif"
        Case ".zip"
            ico="f_rar.gif"
        Case ".htm"
            ico="f_htm.gif"
        Case ".html"
            ico="f_htm.gif"
        Case ".shtml"
            ico="f_htm.gif"
        Case ".asp"
            ico="f_asp.gif"
        Case ".xml"
            ico="f_asp.gif"
        Case ".jsp"
            ico="f_asp.gif"
        Case ".php"
            ico="f_asp.gif"
        Case ".css"
            ico="f_asp.gif"
        Case ".js"
            ico="f_asp.gif"
        Case ".asf"
            ico="f_media.gif"
        Case ".wmv"
            ico="f_media.gif"
        Case ".mdb"
            ico="f_mdb.gif"
        Case ".exe"
            ico="f_exe.gif"
        Case ".com"
            ico="f_exe.gif"
        Case ".bat"
            ico="f_exe.gif"
        Case ".swf"
            ico="f_swf.gif"
        Case ".fla"
            ico="f_swf.gif"
        Case ".rm"
            ico="f_rm.gif"
        Case ".dll"
            ico="f_dll.gif"
        Case ".sys"
            ico="f_dll.gif"
        Case ".ocx"
            ico="f_ocx.gif"
        Case ".ini"
            ico="f_ini.gif"
        Case ".dbx"
            ico="f_dbx.gif"
        Case ".cat"
            ico="f_cat.gif"
        Case ".pdf"
            ico="f_pdf.gif"
        Case ".hlp"
            ico="f_hlp.gif"
        Case ".htt"
            ico="f_htt.gif"
        Case ".png"
            ico="f_png.gif"
        Case ".chm"
            ico="f_chm.gif"
        Case ".nfo"
            ico="f_nfo.gif"
        Case ".reg"
            ico="f_reg.gif"
        Case ".key"
            ico="f_reg.gif"
        Case ".cpp"
            ico="f_cpp.gif"
        Case ".h"
            ico="f_h.gif"
        Case ".frm"
            ico="f_frm.gif"
        Case ".bas"
            ico="f_bas.gif"
        Case ".ctl"
            ico="f_ctl.gif"
        Case ".vbg"
            ico="f_vbg.gif"
        Case ".vbp"
            ico="f_vbp.gif"
        Case else:
            ico="UnKnow.gif"
    End Select
    FileIco=ico

End Function


'########ȡ���ļ���չ��
'����ֵ�磺".exe"��".gif"
Function GetEx(fileName)
	GetEx="."&mid(fileName,InStrRev(fileName, ".")+1)
End Function

'#######��ø�Ŀ¼
Function GetParent(strPath)
    If strPath<>"" Then
        Dim I
        For I=Len(strPath) To 1 Step -1
            If Mid(strPath,I,1)="/" Then
                GetParent=Left(strPath,I-1)
                Exit Function
            End If
        Next
    Else
        GetParent=strPath
    End If
End Function

'#####ɾ��
Function DelItem()
    'If Not SysAdmin.ManageFiles Then
    '    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
    '    Response.End()
    'End If

    Dim ItemList
    DelItem=true
    If Trim(Request("Item"))="" Then
        DelItem=false
        Exit Function
    End If
    ItemList=Split(Trim(Request("Item")),",",-1,1)
    on error resume next
    err.clear
    Dim I,ItemPath
    For I=0 To Ubound(ItemList)
        ItemPath=Server.MapPath(CurrentPath&"/"&Trim(Left(ItemList(I),Len(ItemList(I))-1)))
        '�����ļ�/�ļ��н���ɾ��
        If CBool(Right(ItemList(I),1)) Then
            Fso.DeleteFolder(ItemPath)
        Else
            Fso.DeleteFile(ItemPath)
        End If
        If err.Number<>0 then        
            DelItem=false
        End If
    Next
End Function

'#####�ƶ�
Function MoveItem()
    'If Not SysAdmin.ManageFiles Then
    '    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
    '    Response.End()
    'End If
    Dim DirectionPath
        DirectionPath=""
    Dim Parameter
        Parameter=Request("Parameter")
    If Parameter="/" Then
        DirectionPath=Server.MapPath(DirectoryRoot)&"\"
    Else
        If Left(Parameter,1)="/" Then
            DirectionPath=Server.MapPath(DirectoryRoot&Parameter)&"\"
        Else
            DirectionPath=Server.MapPath(CurrentPath&"/"&Parameter)&"\"
        End If
    End If
    If Not Fso.FolderExists(DirectionPath) Then
        Set Fso=Nothing
        Response.Write("<script>alert(""<����ʧ��>\nĿ¼��"&Parameter&" ������\n�������"");window.history.back();</script>")
        Response.End()
    End If
    MoveItem=true
    If Trim(Request("Item"))="" Then
        DelItem=false
        Exit Function
    End If
    Dim ItemList
        ItemList=Split(Trim(Request("Item")),",",-1,1)
    on error resume next
    err.clear    
    Dim I,ItemPath
    For I=0 To Ubound(ItemList)
        ItemPath=Server.MapPath(CurrentPath&"/"&Trim(Left(ItemList(I),Len(ItemList(I))-1)))
        '�����ļ�/�ļ��н���ɾ��
        If CBool(Right(ItemList(I),1)) Then
            Fso.MoveFolder ItemPath,DirectionPath
        Else
            Fso.MoveFile ItemPath,DirectionPath
        End If
        If err.Number<>0 then        
            MoveItem=false
        End If
    Next
End Function

'#######
Function CreateFolder()
    'If Not SysAdmin.ManageFiles Then
    '    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
    '    Response.End()
    'End If

    Dim Title
    Title=FilterPath(Trim(Request("Title")))
    If Title="" Then
        CreateFolder=false
        Exit Function
    End If
    Dim FolderPath
    on error resume next
    err.clear
    FolderPath=Server.MapPath(CurrentPath&"/"&Title)
    If err.Number<>0 Then
        Response.Write("<script>alert(""<����ʧ��>\nĿ¼�������Ƿ��ַ�"");window.history.back();</script>")
        Response.End
    End If
    If Fso.FolderExists(FolderPath) Then
        Response.Write("<script>alert(""<����ʧ��>\n����ͬ��Ŀ¼"");window.history.back();</script>")
        Response.End
    Else
        Fso.CreateFolder(FolderPath)
        CreateFolder=true
    End If
End Function
%>