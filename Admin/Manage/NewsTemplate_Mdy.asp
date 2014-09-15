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
    var kwList=new Array("资源记录号","资源标题","资源内容","作者","资源来源","关键词","责任编辑","资源小图","资源大图","资源内容摘要","添加时间","修改时间","点击","分类名称","分类别名","分类主页地址","相关资源列表","发表评论链接","评论数目");
    var kwConList=new Array("Id","Title","Content","Author","From","KeyWord","Editor","SmallImg","BigImg","ShortContent","AddTime","UpTime","Count","ClassTitle","ClassTitle2","ClassUrl","ConnectNewsList","Comment","CommentCount");
    for(var i=0;i<kwList.length;i++)
    {
        document.write("<span style=\"cursor:hand\" onclick=\"prompt('系统将自动替换以下关键词为相应的内容,\\n请复制到模板当中','$"+kwConList[i]+"$')\">["+kwList[i]+"]</span> ");
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
        Response.Write("记录未找到")
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
      <td colspan="2" class="BarTitleBg">编辑资源模板</td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">ID:</td>
      <td width="75%" bgcolor="#FFFFFF"> 
        <input name="Id" type="text" class="Input" id="Id2"  value="<%=Id%>" size="4" readonly="true"></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">模板名称:</td>
      <td bgcolor="#FFFFFF"> 
        <input name="Title" type="text" class="Input" id="Title"  value="<%=Title%>" size="40"></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">模板内容:</td>
      <td bgcolor="#f6f6f6"><span style="cursor:hand;color:blue" Id="ShowHiddenHtmlEdit" Title="显示/隐藏" onClick="if(trHtmlEditContent.style.display==''){trHtmlEditContent.style.display='none';ShowHiddenHtmlEdit.innerHTML='[显示]'}else{trHtmlEditContent.style.display='';ShowHiddenHtmlEdit.innerHTML='[隐藏]'}">[显示]</span>&nbsp;<span style="cursor:hand" Title="编辑区加高" onClick="trHtmlEditContent.style.height=1000;">[加高]</span>&nbsp;<span style="cursor:hand" Title="编辑区默认高度" onClick="trHtmlEditContent.style.height=400;">[默认]</span>&nbsp;<span style="cursor:hand" Title="编辑区减低" onClick="trHtmlEditContent.style.height=200;">[减低]</span></td>
    </tr>
    <tr> 
        <td colspan="2" bgcolor="#ffffff"><font color="#0000FF">系统关键替换词:</font> 
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
        alert("请输入[模板名称]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("请输入[模板内容]");
		obj.Content.focus();		
        return false;
    }
    return true;    
}
</script>
      </td>
      <td bgcolor="#FFFFFF"><input name="Submit" type="submit" class="button01-out" value="确  定"> 
        <input name="Submit2" type="reset" class="button01-out" value="还  原"> 
        <input name="Submit3" type="button" class="button01-out" value="返  回" onclick="window.history.back();"> 
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="right" bgcolor="#FFFFFF"> 
        <script>
function DelReco(id){
    if(confirm("你确定删除吗？")){
        window.location="?Work=DelReco&Id="+id;
    }
}
</script>
        <input name="Submit5" type="button" class="button01-out" onclick="DelReco('<%=Id%>')" value="删  除">
      </td>
    </tr>
  </table>
</form>
<%End Sub%>
<%Sub AddReco()%>
<form name="form2" method="post" action="?Work=SaveAddReco" onSubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">添加资源模板</td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">模板名称:</td>
      <td width="75%" bgcolor="#FFFFFF"> 
        <input name="Title" type="text" class="Input" id="Title2" size="40" ></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">模板内容:</td>
      <td bgcolor="#FFFFFF"><span style="cursor:hand;color:blue" Id="ShowHiddenHtmlEdit" Title="显示/隐藏" onClick="if(trHtmlEditContent.style.display==''){trHtmlEditContent.style.display='none';ShowHiddenHtmlEdit.innerHTML='[显示]'}else{trHtmlEditContent.style.display='';ShowHiddenHtmlEdit.innerHTML='[隐藏]'}">[显示]</span>&nbsp;<span style="cursor:hand" Title="编辑区加高" onClick="trHtmlEditContent.style.height=1000;">[加高]</span>&nbsp;<span style="cursor:hand" Title="编辑区默认高度" onClick="trHtmlEditContent.style.height=400;">[默认]</span>&nbsp;<span style="cursor:hand" Title="编辑区减低" onClick="trHtmlEditContent.style.height=200;">[减低]</span></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
        <td colspan="2"><font color="#0000FF">系统关键替换词:</font> 
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
        alert("请输入[模板名称]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("请输入[模板内容]");
		obj.Content.focus();
        return false;
    }
    return true;    
}
</script>
      </td>
      <td bgcolor="#FFFFFF"><input name="Submit4" type="submit" class="button01-out" value="确  定"> 
        <input name="Submit22" type="reset" class="button01-out" value="还  原"> 
        <input name="Submit32" type="button" class="button01-out" value="返  回" onclick="window.history.back();"> 
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
        LogClass.AddLog(SysAdmin.AdminTitle & "试图修改分类模板(Id:"&Request("Id")&")，权限不足")
        Set LogClass=Nothing
        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
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

    LogClass.AddLog(SysAdmin.AdminTitle & "修改资源模板,模板Id:" & Request("Id"))
    Set LogClass=Nothing

    Response.Redirect("NewsTemplate_List.asp")
End Sub

Sub DelReco()
    Dim LogClass
    Set LogClass=New Tkl_LogClass
    If Not SysAdmin.ChangeNewsTemplate Then
        LogClass.AddLog(SysAdmin.AdminTitle & "试图删除分类模板(Id:"&Request("Id")&")，权限不足")
        Set LogClass=Nothing
        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Sql
    Sql="Delete From News_Template Where Id=" & CLng(Request("Id"))
    Conn.ExeCute(Sql)

    LogClass.AddLog(SysAdmin.AdminTitle & "删除资源模板,模板Id:" & CLng(Request("Id")))
    Set LogClass=Nothing

    Response.Redirect("NewsTemplate_List.asp")
End Sub

Sub SaveAddReco()
    If Not SysAdmin.ChangeNewsTemplate Then
        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
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