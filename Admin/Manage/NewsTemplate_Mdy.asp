<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Config.asp" -->
<!-- #include file="Include/Tkl_StringClass.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not SysAdmin.Logined Then
    Response.Redirect("Login.asp")
End If

Dim CFun
Set CFun=New Tkl_StringClass
%>
<html>
<head>
<title>NewsTemplate_Mdy</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--
function TemplateKeyWord()
{
    var kwList=new Array("��Դ��¼��","��Դ����","��Դ����","����","��Դ��Դ","�ؼ���","���α༭","��ԴСͼ","��Դ��ͼ","��Դ����ժҪ","���ʱ��","�޸�ʱ��","���","��������","�������","������ҳ��ַ","�����Դ�б�","������������","������Ŀ");
    var kwConList=new Array("Id","Title","Content","Author","From","KeyWord","Editor","SmallImg","BigImg","ShortContent","AddTime","UpTime","Count","ClassTitle","ClassTitle2","ClassUrl","ConnectNewsList","Comment","CommentCount");
    for(var i=0;i<kwList.length;i++)
    {
        document.write("<span style=\"cursor:hand\" onclick=\"prompt('ϵͳ���Զ��滻���¹ؼ���Ϊ��Ӧ������,\\n�븴�Ƶ�ģ�嵱��','$"+kwConList[i]+"$')\">["+kwList[i]+"]</span> ");
    }
}
//-->
</SCRIPT>
</head>
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
    Case Else
        Call MdyReco()
End Select
%>
<body bgcolor="#FFFFFF">
<%
Sub MdyReco()
    Dim Rs
    Set Rs=Conn.ExeCute("Select * From News_Template Where Id=" & CLng(Request("Id")))
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("��¼δ�ҵ�")
        Response.End
    End If
    Dim Id,Title,Content
    Id=Rs("Id")
    Title=Rs("Title")
    Content=Rs("Content")
    Rs.Close
    Set Rs=Nothing
%>
<form name="form1" method="post" action="?Work=SaveMdy" onSubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr align="center"> 
      <td colspan="2" class="BarTitleBg">�༭��Դģ��</td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">ID:</td>
      <td width="75%" bgcolor="#FFFFFF"> 
        <input name="Id" type="text" class="Input" id="Id2"  value="<%=Id%>" size="4" readonly="true"></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">ģ������:</td>
      <td bgcolor="#FFFFFF"> 
        <input name="Title" type="text" class="Input" id="Title"  value="<%=Title%>" size="40"></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">ģ������:</td>
      <td bgcolor="#f6f6f6"><span style="cursor:hand;color:blue" Id="ShowHiddenHtmlEdit" Title="��ʾ/����" onClick="if(trHtmlEditContent.style.display==''){trHtmlEditContent.style.display='none';ShowHiddenHtmlEdit.innerHTML='[��ʾ]'}else{trHtmlEditContent.style.display='';ShowHiddenHtmlEdit.innerHTML='[����]'}">[��ʾ]</span>&nbsp;<span style="cursor:hand" Title="�༭���Ӹ�" onClick="trHtmlEditContent.style.height=1000;">[�Ӹ�]</span>&nbsp;<span style="cursor:hand" Title="�༭��Ĭ�ϸ߶�" onClick="trHtmlEditContent.style.height=400;">[Ĭ��]</span>&nbsp;<span style="cursor:hand" Title="�༭������" onClick="trHtmlEditContent.style.height=200;">[����]</span></td>
    </tr>
    <tr> 
        <td colspan="2" bgcolor="#ffffff"><font color="#0000FF">ϵͳ�ؼ��滻��:</font> 
        <script language="JavaScript" type="text/JavaScript">TemplateKeyWord();</script>
        </td>
    </tr>
    <tr valign="top"> 
      <td height="400" colspan="2" bgcolor="#FFFFFF" Id="trHtmlEditContent" style="display:"><textarea name="Content" wrap="OFF" class="Input" id="textarea2" style="width=100%;height=100%"><%=CFun.HTMLEncode2(Content)%></textarea></td>
    </tr>
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkMdyReco(obj){
    if(obj.Title.value==""){
        alert("������[ģ������]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("������[ģ������]");
		obj.Content.focus();		
        return false;
    }
    return true;    
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
function DelReco(id){
    if(confirm("��ȷ��ɾ����")){
        window.location="?Work=DelReco&Id="+id;
    }
}
</script>
        <input name="Submit5" type="button" class="button01-out" onclick="DelReco('<%=Id%>')" value="ɾ  ��">
      </td>
    </tr>
  </table>
</form>
<%End Sub%>
<%Sub AddReco()%>
<form name="form2" method="post" action="?Work=SaveAddReco" onSubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">�����Դģ��</td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">ģ������:</td>
      <td width="75%" bgcolor="#FFFFFF"> 
        <input name="Title" type="text" class="Input" id="Title2" size="40" ></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">ģ������:</td>
      <td bgcolor="#FFFFFF"><span style="cursor:hand;color:blue" Id="ShowHiddenHtmlEdit" Title="��ʾ/����" onClick="if(trHtmlEditContent.style.display==''){trHtmlEditContent.style.display='none';ShowHiddenHtmlEdit.innerHTML='[��ʾ]'}else{trHtmlEditContent.style.display='';ShowHiddenHtmlEdit.innerHTML='[����]'}">[��ʾ]</span>&nbsp;<span style="cursor:hand" Title="�༭���Ӹ�" onClick="trHtmlEditContent.style.height=1000;">[�Ӹ�]</span>&nbsp;<span style="cursor:hand" Title="�༭��Ĭ�ϸ߶�" onClick="trHtmlEditContent.style.height=400;">[Ĭ��]</span>&nbsp;<span style="cursor:hand" Title="�༭������" onClick="trHtmlEditContent.style.height=200;">[����]</span></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
        <td colspan="2"><font color="#0000FF">ϵͳ�ؼ��滻��:</font> 
        <script language="JavaScript" type="text/JavaScript">TemplateKeyWord();</script>
        </td>
    </tr>
    <tr valign="top"> 
      <td height="400" colspan="2" bgcolor="#FFFFFF" Id="trHtmlEditContent" style="display:"> 
        <textarea name="Content" wrap="OFF" class="Input" id="textarea2" style="width=100%;height=100%"></textarea>
        
      </td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkAddReco(obj){
    if(obj.Title.value==""){
        alert("������[ģ������]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("������[ģ������]");
		obj.Content.focus();
        return false;
    }
    return true;    
}
</script>
      </td>
      <td bgcolor="#FFFFFF"><input name="Submit4" type="submit" class="button01-out" value="ȷ  ��"> 
        <input name="Submit22" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit32" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"> 
      </td>
    </tr>
  </table>
</form>
<%End Sub%>
</body>
</html>
<%
Sub SaveMdy()
    Dim LogClass
    Set LogClass=New Tkl_LogClass
    If Not SysAdmin.ChangeNewsTemplate Then
        LogClass.AddLog(SysAdmin.AdminTitle & "��ͼ�޸ķ���ģ��(Id:"&Request("Id")&")��Ȩ�޲���")
        Set LogClass=Nothing
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Sql
        Sql="Select * From News_Template Where Id=" & CLng(Request("Id"))
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    If Not (Rs.Eof And Rs.Bof) Then
        Rs("Title")= Trim(Request("Title"))
        Rs("Content")= Trim(Request("Content"))
        Rs("upTime")= Now
        Rs.Update
    End If
    Rs.Close
    Set Rs=Nothing

    LogClass.AddLog(SysAdmin.AdminTitle & "�޸���Դģ��,ģ��Id:" & Request("Id"))
    Set LogClass=Nothing

    Response.Redirect("NewsTemplate_List.asp")
End Sub

Sub DelReco()
    Dim LogClass
    Set LogClass=New Tkl_LogClass
    If Not SysAdmin.ChangeNewsTemplate Then
        LogClass.AddLog(SysAdmin.AdminTitle & "��ͼɾ������ģ��(Id:"&Request("Id")&")��Ȩ�޲���")
        Set LogClass=Nothing
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Sql
    Sql="Delete From News_Template Where Id=" & CLng(Request("Id"))
    Conn.ExeCute(Sql)

    LogClass.AddLog(SysAdmin.AdminTitle & "ɾ����Դģ��,ģ��Id:" & CLng(Request("Id")))
    Set LogClass=Nothing

    Response.Redirect("NewsTemplate_List.asp")
End Sub

Sub SaveAddReco()
    If Not SysAdmin.ChangeNewsTemplate Then
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Sql
        Sql="Select Top 1 * From News_Template Order By ID DESC"
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    Rs.AddNew
    Rs("Title")= Trim(Request("Title"))
    Rs("Content")= Trim(Request("Content"))
    Rs("upTime")= Now
    Rs("AddTime")= Now
    Rs.Update
    Rs.Close
    Set Rs=Nothing
    Response.Redirect("NewsTemplate_List.asp")
End Sub
%>