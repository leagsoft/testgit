<!--#include file="Include/Conn.asp" -->
<!-- #include file="Include/ClassList_Fun.asp" -->
<!-- #include file="Include/Tkl_StringClass.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not SysAdmin.Logined Then
'    Response.Redirect("Login.asp")
'End If

%>
<html>
<head>
<title>�����޸�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>

<body bgcolor="#FFFFFF">
<%
Select Case Request("Work")
    Case "SaveMdy"
        Call SaveMdy()
    Case "DelReco"
        Call DelReco()
    Case "AddReco"
        Call AddReco()
    Case "SaveAddReco"
        Call SaveAddReco()
	Case "ClearFolder"
		Call ClearFolder()
	Case "DirectoryInfo"
		Call DirectoryInfo()		
    Case Else
        Call MdyReco()
End Select
%>
<%
Sub MdyReco()
    Dim Rs,Rs2
    Dim Sql,Sql2
        Sql="Select CL.Id,CL.Title,CL.Title2,CL.Directory,CL.ClassUrl,CL.Template,CL.Parent,CL.UpTime,NT.Title As TemplateTitle From ClassList CL LEFT JOIN  News_Template NT ON NT.Id=CL.Template Where CL.Id=" & Request("Id")
    Set Rs=Conn.ExeCute(Sql)
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("��¼δ�ҵ�")
        Response.End
    End If
    Dim Id,Title,Title2,Parent,Directory,ClassUrl,Template,TemplateTitle
    Id=Rs("Id")
    Title=Rs("Title")
    Title2=Rs("Title2")
    Parent=Rs("Parent")
    Directory=Rs("Directory")
    ClassUrl=Rs("ClassUrl")
    Template=Rs("Template")
    TemplateTitle=Rs("TemplateTitle")
    Rs.Close
    Set Rs=Nothing
%>
<form name="form1" method="post" action="?Work=SaveMdy" onSubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">�༭��Դ����</td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">�����ID:</td>
      <td width="75%" bgcolor="#FFFFFF"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="34%"> <input name="Parent" type="text" class="Input" id="Parent3" value="<%=Parent%>" size="4"></td>
            <td width="66%"><font color="#666666">����ת��[����]������������,�����ظ���</font></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">��ϸλ��:</td>
      <td bgcolor="#FFFFFF"> <%=GetClassPath2(SysAdmin.AdminTopClassId,Id,"Class_List.asp?")%> <input name="Id" type="hidden" id="Id" value="<%=Request("Id")%>"> 
      </td>
    </tr>
    <tr>
      <td align="right" class="BarTitle">��������:</td>
      <td bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title3" value="<%=Title%>" size="40"> 
      </td>
    </tr>
    <tr>
      <td align="right" class="BarTitle">�������:</td>
      <td bgcolor="#FFFFFF"> <input name="Title2" type="text" class="Input" id="Title4" value="<%=Title2%>" size="40" onfocus="if(Title.value!='' && this.value==''){this.value=Title.value}"> 
      </td>
    </tr>
    <!--<tr> 
      <td align="right" class="BarTitle">����Ŀ¼:</td>
      <td bgcolor="#FFFFFF"> <input name="Directory" type="text" class="Input" id="Directory" value="<%=Directory%>" size="40">
        <input name="Submit33" type="button" class="button01-out" value="�� ��" onclick="ClearFolder(<%=Id%>)">
      </td>
    </tr>
    <tr>
      <td align="right" class="BarTitle">������ҳ��ַ:</td>
      <td bgcolor="#FFFFFF"> <input name="ClassUrl" type="text" class="Input" id="ClassUrl" value="<%=ClassUrl%>" size="40"> 
      </td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">ʹ��ģ��:</td>
      <td bgcolor="#FFFFFF"> 
        <select name="Template" id="Template1">
        <option value="" selected>��Դģ��</option>        
<%
            Sql2="SELECT * From News_Template Order By upTime DESC"
            Set Rs2=Conn.ExeCute(Sql2)
            While Not Rs2.Eof
%>
          <option value="<%=Rs2("Id")%>" <%If Template=Rs2("Id") Then Response.Write("Selected") End If%>><%=Rs2("Title")%></option>
<%
                Rs2.MoveNext
            Wend
            Rs2.Close
            Set Rs2=Nothing
%>
        </select> </td>
    </tr>-->

    <tr> 
      <td align="right" class="BarTitle"> 
        <script>
function ClearFolder(ClassId)
{
	if(confirm("<����>\n��ȷ������Ŀ¼ɾ���������ڲ��������ļ���ͬʱ��ɾ����"))
	{
		window.location="?Work=ClearFolder&ClassId="+ClassId;
	}
}		
function checkMdyReco(obj){
    if(obj.Title.value==""){
        alert("������[��������]");
        obj.Title.focus();
        return false;
    }
    if(obj.Title2.value==""){
        alert("������[�������]");
        obj.Title2.focus();
        return false;
    };
    //if(obj.Directory.value==""){
    //    alert("������[����Ŀ¼]");
    //    obj.Directory.focus();
    //    return false;
   // }else{
   //     if(obj.Directory.value.search(/^[\/|\\]/i)==-1){
   //         alert("[����Ŀ¼]�����ԡ�/����Ŀ¼��ʼ");
   //         obj.Directory.focus();
   //         return false;
   //     }else{
   //         obj.Directory.value=obj.Directory.value.replace(/[\/|\\]{1,}$/i,"");
   //     }
   // }
   // if(obj.ClassUrl.value==""){
   //     alert("������[������ҳ���ַ]");
   //     obj.ClassUrl.focus();
   //     return false;
   // }    
   // if(obj.Template.value==""){
    //    alert("<����>\n�˷���û������ģ��")
    //    return false
   // };
    return true
}
</script>
</td>
      <td bgcolor="#FFFFFF"><input name="Submit" type="submit" class="button01-out" value="ȷ  ��"> 
        <input name="Submit2" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit3" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"> 
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="right" bgcolor="#FFFFFF">
<script>
function DelReco(id,Parent){
	var url="?Work=DelReco&Parent=" + Parent + "&Id="+id;
    if(confirm("��ȷ��ɾ����")){
        window.location=url
    }
}
</script>        <label for="DelResource"></label>
        <input name="Submit5" type="button" class="button01-out" onclick="DelReco('<%=Id%>','<%=Parent%>')" value="ɾ  ��">
      </td>
    </tr>
  </table>
</form>
<%End Sub%>
<%
Sub AddReco()
    Dim Rs
    Dim Sql
%>
<form name="form2" method="post" action="?Work=SaveAddReco" onSubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" bgcolor="#CCCCCC" class="BarTitleBg">�����Դ����</td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">��ϸλ��:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Parent" type="hidden" id="Parent" value="<%=Request("Parent")%>"> 
        <%=GetClassPath2(SysAdmin.AdminTopClassId,Request("Parent"),"Class_List.asp?")%> </td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">��������:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title5" size="40"> 
      </td>
    </tr>
    <tr>
      <td align="right" class="BarTitle">�������:</td>
      <td bgcolor="#FFFFFF"> <input name="Title2" type="text" class="Input" id="Title4" value="" size="40" onfocus="if(Title.value!='' && this.value==''){this.value=Title.value}"> 
      </td>
    </tr>
    <!--<tr> 
      <td align="right" class="BarTitle">ָ��Ŀ¼:</td>
      <td bgcolor="#FFFFFF"> <input name="Directory" type="text" class="Input" id="Directory" value="" size="40">
      </td>
    </tr>
    <tr>
      <td align="right" class="BarTitle">������ҳ��ַ:</td>
      <td bgcolor="#FFFFFF"> <input name="ClassUrl" type="text" class="Input" id="ClassUrl" value="http://" size="40"> 
      </td>
    </tr>-->
    <!--<tr> 
      <td align="right" class="BarTitle">ʹ��ģ��:</td>
      <td bgcolor="#FFFFFF">
        <select name="Template" id="Template2">
        <option value="" selected>��Դģ��</option>
<%
            Sql="SELECT * From News_Template Order By upTime DESC"
            Set Rs=Conn.ExeCute(Sql)
            While Not Rs.Eof
%>
          <option value="<%=Rs("Id")%>"><%=Rs("Title")%></option>
<%
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs=Nothing
%>
        </select>
      </td>
    </tr>-->
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkAddReco(obj){
    if(obj.Title.value==""){
        alert("������[��������]");
        obj.Title.focus();
        return false;
    }
    if(obj.Title2.value==""){
        alert("������[�������]");
        obj.Title2.focus();
        return false;
    };
   //if(obj.Directory.value==""){
   //     alert("������[����Ŀ¼]");
   //     obj.Directory.focus();
   //     return false;
   // }else{
   //     if(obj.Directory.value.search(/^[\/|\\]/i)==-1){
   //         alert("[����Ŀ¼]�����ԡ�/����Ŀ¼��ʼ");
   //         obj.Directory.focus();
   //         return false;
   //     }else{
   //         obj.Directory.value=obj.Directory.value.replace(/[\/|\\]{1,}$/i,"");
   //     }
   // }
   // if(obj.ClassUrl.value==""){
   //     alert("������[������ҳ���ַ]");
   //     obj.ClassUrl.focus();
   //     return false;
   // }
   // if(obj.Template.value==""){
   //     alert("<����>\n�˷���û������ģ��")
   //     return false;
   // };
    return true;
}
</script>
</td>
      <td bgcolor="#FFFFFF"><input name="Submit4" type="submit" class="button01-out" value="ȷ  ��"> 
        <input name="Submit22" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit32" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
    </tr>
  </table>
</form>
<%End Sub%>
<%
Sub DirectoryInfo()
    Dim Rs
    Dim Sql
        Sql="Select Id,Title,Title2,Directory From ClassList Where Id=" & CLng(Request("Id"))
    Set Rs=Conn.ExeCute(Sql)
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("��¼δ�ҵ�")
        Response.End
    End If
    Dim Id,Title,Title2,Directory
    Id=Rs("Id")
    Title=Rs("Title")
    Title2=Rs("Title2")
    Directory=Rs("Directory")
    Rs.Close
    Set Rs=Nothing
	
	Dim Fso,Fol
	Set Fso = Server.CreateObject(FsoObjectStr)
	If Not Fso.FolderExists(Server.MapPath(Directory)) Then
		Set Fso=Nothing	
		Response.Write("<script>alert(""<����ʧ��>\n��ǰ���������Ŀ¼������"& SoftCopyright_Script &""");window.history.back();</script>")
		Response.End()
	End If
	Set Fol = Fso.GetFolder(Server.MapPath(Directory))
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td colspan="2" align="center" class="BarTitleBg">&quot;<%=Title%>&quot;��������Ŀ¼״̬��Ϣ</td>
  </tr>
  <tr> 
    <td width="25%" align="right" class="BarTitle">Ŀ¼�߼�λ��:</td>
    <td width="75%" bgcolor="#FFFFFF"><%=Directory%></td>
  </tr>
  <tr> 
    <td align="right" class="BarTitle"> 
      <p>Ŀ¼����λ��:</p>
    </td>
    <td width="75%" bgcolor="#FFFFFF"><%=Fol.Path%></td>
  </tr>
  <tr> 
    <td align="right" class="BarTitle">Ŀ¼��С:</td>
    <td bgcolor="#FFFFFF"><%=FormatNumber(Fol.Size)%> �ֽ�</td>
  </tr>
  <tr> 
    <td align="right" class="BarTitle">��Ŀ¼��:</td>
    <td bgcolor="#FFFFFF"><%=Fol.SubFolders.Count%></td>
  </tr>
  <tr> 
    <td align="right" class="BarTitle">Ŀ¼����ʱ��:</td>
    <td bgcolor="#FFFFFF"><%=Fol.DateCreated%></td>
  </tr>
  <tr> 
    <td align="right" class="BarTitle">Ŀ¼���洢ʱ��:</td>
    <td bgcolor="#FFFFFF"><%=Fol.DateLastAccessed%></td>
  </tr>
  <tr> 
    <td align="right" class="BarTitle">Ŀ¼����޸�ʱ��:</td>
    <td bgcolor="#FFFFFF"><%=Fol.DateLastModified%></td>
  </tr>
  <tr>
    <td align="right" class="BarTitle">&nbsp;</td>
    <td bgcolor="#FFFFFF">
      <input name="Submit34" type="button" class="button01-out" value="��  ��" onclick="window.history.back();">
    </td>
  </tr>
</table>
<%
End Sub
%>
<script language="JavaScript" type="text/JavaScript">
document.write("<iframe style=\"display:none\" id=\"ActionFrame\"></iframe>");
function ShowControlPane(obj)
{
    if(obj.style.display=='')
    {
        obj.style.display='none'
    }else{
        obj.style.display=''
    }
}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display='';window.scrollTo(window.pageXOffset,2000);}else{HelpTab.style.display='none'}">&nbsp;<img src="Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:none"> 
    <td width="2%">&nbsp;</td>
    <td width="98%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td>������Ŀ¼</td>
        </tr>
        <tr>
          <td>����ָ����Դ������Ŀ¼������վ�������Ŀ¼������Χ�ڣ�������Ŀ¼�����ԡ�/��վ���Ŀ¼Ϊ������ã��磺/new��</td>
        </tr> 
      </table>
      <a name="Help"></a></td>
  </tr>
</table>
</body>
</html>
<%
Sub SaveMdy()
    'If Not SysAdmin.ChangeCommentList Then

     '   Dim LogClass
     '   Set LogClass=New Tkl_LogClass
     '   LogClass.AddLog(SysAdmin.AdminTitle & "��ͼ�޸���Դ����(Id:"&Request("Id")&")��Ȩ�޲���")
     '   Set LogClass=Nothing

    '    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
    '    Response.End()
    'End If
    
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
    Dim Sql
        Sql="Select Top 1 * From ClassList Where Id=" & CLng(Request("Id"))
    Rs.Open Sql,Conn,1,3
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("<script>alert('<����ʧ��>\n��¼������"& SoftCopyright_Script &"');window.history.back();</script>")
        Response.End()
    End If
    Rs("Title")=Request("Title")
    Rs("Title2")=Request("Title2")
    Rs("Parent")=Request("Parent")
	'************* Modify By BennyLiu:20040311 ***************
    'Rs("Directory")=Request("Directory")
    'Rs("ClassUrl")=Request("ClassUrl")
    'Rs("Template")=Request("Template")
    '******************* End Modify **************************
    Rs("upTime")=Now
    Rs.Update
    Response.Redirect("Class_List.asp?Parent=" & Request("Parent"))
End Sub

Sub DelReco()
    'If Not SysAdmin.ChangeCommentList Then

        'Dim LogClass
        'Set LogClass=New Tkl_LogClass
        'LogClass.AddLog(SysAdmin.AdminTitle & "��ͼɾ����Դ����(Id:"&Request("Id")&")��Ȩ�޲���")
        'Set LogClass=Nothing

    '    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
    '    Response.End()
    'End If

    Dim Sql,Rs
	'�ж��Ƿ����»��������ڣ��粻Ϊ����÷��಻��ɾ��
    Sql="Select Count(*) As Num From News Where Class In ("& Request("Id") & AllChildClass(Request("Id")) &")"
    Set Rs=Conn.ExeCute(Sql)
    If Rs("Num")>=1 Then
        Response.Write("<script>alert('<����ʧ��>\n���»���[��Դ]����Щ�޷�ɾ����\n����ɾ�����µ���Դ����ɾ���˷��࣡"& SoftCopyright_Script &"');window.history.back();</script>")
        Rs.Close
		Response.End
    End If
    Rs.Close

    Sql="Select Count(*) As Num From ClassList Where Parent="&Request("Id")
    Set Rs=Conn.ExeCute(Sql)
    If Rs("Num")>=1 Then
        Response.Write("<script>alert('<����ʧ��>\n���»���[����]���޷�ɾ����"& SoftCopyright_Script &"');window.history.back();</script>")
        Rs.Close
		Response.End
	End If
	Rs.Close

	'�Ƿ���ĳһ��ɫ�ģ����ƽ��ɲ鿴�ķ��࣢����ʹ�ø����
    'Sql="Select Count(*) As Num From Admin_Role Where ClassId="&Request("Id")
    'Set Rs=Conn.ExeCute(Sql)
    'If Rs("Num")>=1 Then
    '    Response.Write("<script>alert('<����ʧ��>\n���»���[����Ա��ɫ�ģ����ƽ��ɲ鿴�ķ��ࣧʹ�ø����]������޷�ɾ����"& SoftCopyright_Script &"');window.history.back();</script>")
    '    Rs.Close
	'	Response.End
    'End If
	'Rs.Close

	'ɾ������
	Sql="Delete From ClassList Where Id=" & Request("Id")
	Conn.ExeCute(Sql)
	Response.Redirect("Class_List.asp?Parent="&Request("Parent"))

End Sub

Sub SaveAddReco()
    'If Not SysAdmin.ChangeCommentList Then
    '    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
    '    Response.End()
    'End If

    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
    Dim Sql
        Sql="Select Top 1 * From ClassList"
    Rs.Open Sql,Conn,1,3
    Rs.AddNew
    Rs("Title")=Request("Title")
    Rs("Title2")=Request("Title2")
    Rs("Parent")=Request("Parent")
   '********************** Modify By BennyLiu:20040311 *************
    'Rs("Directory")=Request("Directory")
    'Rs("ClassUrl")=Request("ClassUrl")
    'Rs("Template")=Request("Template")
   '************************* End Modify ***************************
    Rs("upTime")=Now
    Rs.Update
    Response.Redirect("Class_List.asp?Parent="&Request("Parent"))
End Sub

Sub ClearFolder()
    'If Not SysAdmin.ChangeCommentList Then
    '    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
    '    Response.End()
    'End If
	
	Dim Rs
	Set Rs=Conn.ExeCute("Select Top 1 Directory From ClassList Where Id="&CLng(Request("ClassId")))
	If Rs.Eof And Rs.Bof Then
		Rs.Close
		Set Rs=Nothing
		Response.Write("<script>alert(""<����ʧ��>\n��ǰ�����Ѳ�����"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
	End If
	
	Dim Fso
	Set Fso = Server.CreateObject(FsoObjectStr)
	If Fso.FolderExists(Server.MapPath(Rs("Directory"))) Then
		Fso.DeleteFolder(Server.MapPath(Rs("Directory")))
	End If
	Set Fso=Nothing
	Rs.Close
	Set Rs=Nothing
	Response.Write("<script>alert(""<����ʧ��>\n��ǰ���������Ŀ¼���ڲ������ļ������"& SoftCopyright_Script &""");window.history.back();</script>")
	Response.End()	
End Sub
%>
