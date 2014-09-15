<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#include file="CheckAdmin.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Tkl_TemplateClass.asp" -->
<!--#include file="Include/Tkl_StringClass.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not CBool(SysAdmin.Logined) Then
'    Response.Redirect("Login.asp")
'End If

Dim cFun
Set cFun=New Tkl_StringClass
%>
<html>
<head>
<title>InsertSYS_Mdy.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
<script src="Library/htmlarea/init_htmlarea.js"></script>
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
    Case "InsertSysActive"
        Call InsertSysActive()
    Case "DoInsertSysActive"
        Call DoInsertSysActive()
    Case Else
        Call MdyReco()
End Select
%>
<%
Sub MdyReco()
    Dim Rs
    Set Rs=Conn.ExeCute("Select * From InsertList Where Id=" & Request("Id"))
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("记录未找到")
        Response.End
    End If
    Dim Id,Title,StartElement,EndElement,FileList,Content
    Id=Rs("Id")
    Title=Rs("Title")
    StartElement=Rs("StartElement")
    EndElement=Rs("EndElement")
    FileList=Rs("FileList")
    Content=Rs("Content")
    Rs.Close
    Set Rs=Nothing
%>
<form name="form1" method="post" action="?Work=SaveMdy" onSubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">编辑内容嵌入块</td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">ID:</td>
      <td width="75%" bgcolor="#FFFFFF"> 
        <input name="Id" type="text" class="Input" id="Id2"  value="<%=Id%>" size="4" readonly="true">
      </td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">标题:</td>
      <td width="75%" bgcolor="#FFFFFF"> 
        <input name="Title" type="text" class="Input" id="Title2" size="40" value="<%=Title%>">
      </td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">起始标签:</td>
      <td bgcolor="#FFFFFF"> 
        <input name="StartElement" type="text" class="Input" id="StartElement" size="40" value="<%=cFun.HTMLEncode2(StartElement)%>">
      </td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">结束标签:</td>
      <td bgcolor="#FFFFFF"> 
        <input name="EndElement" type="text" class="Input" id="EndElement" size="40" value="<%=CFun.HTMLEncode2(EndElement)%>">
      </td>
    </tr>
    <tr> 
      <td height="300" colspan="2" valign="top" bgcolor="buttonface"> 
        <textarea name="Content" cols="40" wrap="OFF" id="Content"><%=CFun.HTMLEncode2(Content)%></textarea>
      </td>
    </tr>
    <tr> 
      <td height="85" align="right" valign="top" class="BarTitle">应用文件:<br>
        <font color="#999999">一行一个文件地址</font> </td>
      <td valign="top" bgcolor="#FFFFFF"> 
        <textarea name="FileList" cols="40" rows="5" wrap="OFF" class="Input" id="FileList" style="width:100%"><%=FileList%></textarea>
      </td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkMdyReco(obj){
    if(obj.Title.value==""){
        alert("请输入[标题]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("请输入[显示代码]");
        obj.Content.focus();
        return false;
    }
    return true;
}
</script>
      </td>
      <td bgcolor="#FFFFFF"> 
        <input name="Submit" type="submit" class="button01-out" value="确  定">
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
        <input name="Submit52" type="button" class="button01-out" onclick="DelReco('<%=Id%>')" value="删  除">
      </td>
    </tr>
  </table>
</form>
<script language="javascript1.2">
editor_generate('Content',config);
</script>
<%End Sub%>
<%Sub AddReco()%>
<form name="form2" method="post" action="?Work=SaveAddReco" onSubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">添加内容嵌入块</td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">标题:</td>
      <td width="75%" bgcolor="#FFFFFF"> 
        <input name="Title" type="text" class="Input" id="Title2" size="40" >
      </td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">起始标签:</td>
      <td bgcolor="#FFFFFF"> 
        <input name="StartElement" type="text" class="Input" id="StartElement" size="40" >
      </td>
    </tr>
    <tr> 
      <td width="25%" align="right" class="BarTitle">结束标签:</td>
      <td bgcolor="#FFFFFF"> 
        <input name="EndElement" type="text" class="Input" id="EndElement" size="40" >
      </td>
    </tr>
    <tr> 
      <td height="300" colspan="2" align="right" valign="top" bgcolor="buttonface"> 
        <textarea name="Content" cols="40" wrap="OFF" class="Input" id="Content3" style="width:100%;height:100%"></textarea>
      </td>
    </tr>
    <tr> 
      <td height="85" align="right" valign="top" class="BarTitle">应用文件:<br>
        <font color="#999999">一行一个文件地址</font> </td>
      <td valign="top" bgcolor="#FFFFFF"> 
        <textarea name="FileList" cols="40" rows="5" wrap="OFF" class="Input" id="Content3" style="width:100%"></textarea>
      </td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkAddReco(obj){
    if(obj.Title.value==""){
        alert("请输入[标题]");
        obj.Title.focus();
        return false;
    }
    if(obj.StartElement.value==""){
        alert("请输入[起始标签]");
        obj.StartElement.focus();
        return false;
    }
    if(obj.EndElement.value==""){
        alert("请输入[结束标签]");
        obj.EndElement.focus();
        return false;
    }
    if(obj.FileList.value==""){
        alert("请输入[应用文件列表]");
        obj.FileList.focus();
        return false;
    }
    return true;
}
</script>
      </td>
      <td bgcolor="#FFFFFF"> 
        <input name="Submit4" type="submit" class="button01-out" value="确  定">
        <input name="Submit22" type="reset" class="button01-out" value="还  原">
        <input name="Submit32" type="button" class="button01-out" value="返  回" onclick="window.history.back();">
      </td>
    </tr>
  </table>
</form>
<script language="javascript1.2">
editor_generate('Content',config);
</script>
<%End Sub%>
<%Sub InsertSysActive()%>
<table width="500" height="116" border="0" align="center" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td height="21" align="center" class="BarTitleBg">执行页面内容替换命令(请使用Ctrl\Shift组合键)</td>
  </tr>
  <tr>
    <td height="92" align="center" valign="top" bgcolor="#FFFFFF">
    <form name="form3" method="post" action="?Work=DoInsertSysActive">
        <p> 
          <select name="RecordList" size="25" multiple class="Input" style="width:100%">
<%
    Dim Sql
        Sql="Select Id,Title,upTime From InsertList Order By Id DESC"
    Dim Rs
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
        </p>
        <input name="Submit5" type="submit" class="button01-out" value="确  定">
        <input name="Submit23" type="reset" class="button01-out" value="还  原">
        <input name="Submit33" type="button" class="button01-out" value="返  回" onClick="window.history.back();">
      </form>
      
    </td>
  </tr>
</table>
<%End Sub%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display='';window.scrollTo(window.pageXOffset,2000);}else{HelpTab.style.display='none'}">&nbsp;<img src="Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:none"> 
    <td width="2%">&nbsp;</td>
    <td width="98%" valign="top"> <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>页面内容嵌入功能介绍：</td>
        </tr>
        <tr> 
          <td>　　此功能模块可以帮助管理员对站点页面中的各小块内容进行在线管理及更新成静态文件。其适用的范围如：页面中的小广告、站点通告、版权内容及其它一些页面中的边角内容块<br> 
          </td>
        </tr>
      </table>
      <a name="Help"></a></td>
  </tr>
</table>
</body>
</html>
<%
Sub SaveMdy()
    If Not SysAdmin.InsertSYS Then
        Dim LogClass
        Set LogClass=New Tkl_LogClass
        LogClass.AddLog(SysAdmin.AdminTitle & "试图修改内容替换数据(Id:"&Request("Id")&")，权限不足")
        Set LogClass=Nothing
        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Sql
        Sql="Select * From InsertList Where Id=" & Request("Id")
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    If Not (Rs.Eof And Rs.Bof) Then
        Rs("Title")= Trim(Request("Title"))
        Rs("StartElement")= Trim(Request("StartElement"))
        Rs("EndElement")= Trim(Request("EndElement"))
        Rs("Content")= Trim(Request("Content"))
        Rs("FileList")= Trim(Request("FileList"))
        Rs("upTime")= Now
        Rs.Update
    End If
    Rs.Close
    Set Rs=Nothing
    Response.Redirect("InsertSYS_List.asp")
End Sub

Sub DelReco()
    If Not SysAdmin.InsertSYS Then
        Dim LogClass
        Set LogClass=New Tkl_LogClass
        LogClass.AddLog(SysAdmin.AdminTitle & "试图删除内容替换数据(Id:"&Request("Id")&")，权限不足")
        Set LogClass=Nothing

        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If
    
    Dim Sql
    Sql="Delete From InsertList Where Id=" & Request("Id")
    Conn.ExeCute(Sql)
    Response.Redirect("InsertSYS_List.asp")
End Sub

Sub SaveAddReco()
	If Not SysAdmin.InsertSYS Then
		Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
		Response.End()
	End If

    Dim Sql
        Sql="Select Top 1 * From InsertList"
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    Rs.AddNew
    Rs("Title")= Trim(Request("Title"))
    Rs("StartElement")= Trim(Request("StartElement"))
    Rs("EndElement")= Trim(Request("EndElement"))
    Rs("Content")= Trim(Request("Content"))
    Rs("FileList")= Trim(Request("FileList"))
    Rs("AddTime")= Now
    Rs("upTime")= Now
    Rs.Update
    Rs.Close
    Set Rs=Nothing
    Response.Redirect("InsertSYS_List.asp")
End Sub

Sub DoInsertSysActive()
    Dim LogClass
    Set LogClass=New Tkl_LogClass
    If Not SysAdmin.InsertSYS Then

        LogClass.AddLog(SysAdmin.AdminTitle & "试图执行内容替换，权限不足")
        Set LogClass=Nothing

        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

   On Error Resume Next
    Dim RecordList
        RecordList=Trim(Replace(Request("RecordList")," ",""))
    If RecordList="" Then
        Response.Write("<script>alert(""<操作失败>\n请您选择要执行的[内容替换命令]"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End
    End If
    Dim arrRecordList
        arrRecordList=Split(RecordList,",",-1,1)
    Dim Sql
        Sql=""
    Dim Rs
    Dim arrFileList
    Dim TClass
    Set TClass=New Tkl_TemplateClass

    Dim I,J
    For I=0 To UBound(arrRecordList)
        Sql="Select * From InsertList Where Id=" & arrRecordList(I)
        Set Rs=Conn.ExeCute(Sql)
        If Not(Rs.Eof And Rs.Bof) Then
            arrFileList=Split(Rs("FileList"),vbCrLf,-1)
            For J=0 To UBound(arrFileList)
                With TClass
                    .OpenTemplate(Server.MapPath(arrFileList(J)))
                    .StartElement=Rs("StartElement")
                    .EndElement=Rs("EndElement")
                    .Value=Rs("Content")
                    .ReplaceTemplate()
                    .Save()
                End With
            Next
        End If
        Rs.Close
    Next
    Set TClass=Nothing
    Set Rs=Nothing

    LogClass.AddLog(SysAdmin.AdminTitle & "执行内容替换操作")
    Set LogClass=Nothing

    Response.Write("<script>alert(""<操作成功>\n内容替换完毕"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End
End Sub
%>