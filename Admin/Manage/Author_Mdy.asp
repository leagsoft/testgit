<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If

If Not SysAdmin.ChageAuthor Then
    Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End If
%>
<html>
<head>
<title>From_Mdy</title>
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
    Case Else
        Call MdyReco()
End Select
%>
<%Sub MdyReco()
    Dim Rs
    Set Rs=Conn.ExeCute("Select * From AuthorList Where Id=" & Request("Id"))
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("记录未找到")
        Response.End
    End If
    Dim Id,Title,Sex,Email,Content,BigPhoto
    Id=Rs("Id")
    Title=Rs("Title")
    Sex=Rs("Sex")
    Email=Rs("Email")
    Content=Rs("Content")
    BigPhoto=Rs("BigPhoto")
    Rs.Close
    Set Rs=Nothing
%>
<form name="form1" method="post" action="?Work=SaveMdy" onSubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">编辑作者信息</td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">ID:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Id" type="text" class="Input" id="Id2"  value="<%=Id%>" size="4" readonly="true"></td>
    </tr>
    <tr> 
      <td class="BarTitle">作者姓名（笔名）:</td>
      <td bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title"  value="<%=Title%>" size="40"></td>
    </tr>
    <tr> 
      <td class="BarTitle">性别:</td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Sex" value="0" <%If CLng(Sex)=0 Then Response.Write("checked") End If%>>
        女　 
        <input type="radio" name="Sex" value="1" <%If CLng(Sex)=1 Then Response.Write("checked") End If%>>
        男　 
        <input name="Sex" type="radio" value="2" <%If CLng(Sex)=2 Then Response.Write("checked") End If%>>
        保密</td>
    </tr>
    <tr> 
      <td class="BarTitle">Email:</td>
      <td bgcolor="#FFFFFF"><input name="Email" type="text" class="Input" id="Email"  value="<%=Email%>" size="40"></td>
    </tr>
    <tr>
      <td class="BarTitle">照片:</td>
      <td bgcolor="#FFFFFF"><input name="BigPhoto" type="text" class="Input" id="BigPhoto" size="40" value="<%=BigPhoto%>"></td>
    </tr>
    <tr> 
      <td width="25%" valign="top" class="BarTitle">作者简介:</td>
      <td width="75%" bgcolor="#FFFFFF"><textarea name="Content" cols="60" rows="8" class="Input" id="Content"><%=Content%></textarea></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkMdyReco(obj){
    if(obj.Title.value==""){
        alert("请输入[作者姓名（笔名）]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("请输入[作者简介]");
        obj.Content.focus();
        return false;
    }
    return true;    
}
</script> </td>
      <td bgcolor="#FFFFFF"> <input name="Submit" type="submit" class="button01-out" value="确  定"> 
        <input name="Submit2" type="reset" class="button01-out" value="还  原"> 
        <input name="Submit3" type="button" class="button01-out" value="返  回" onclick="window.history.back();"></td>
    </tr>
    <tr> 
      <td colspan="2" align="right" bgcolor="#FFFFFF"> <script>
function DelReco(id){
    if(confirm("你确定删除吗？")){
        window.location="?Work=DelReco&Id="+id;
    }
}
</script>
        <input name="Submit5" type="button" class="button01-out" onclick="DelReco('<%=Id%>')" value="删 除" <%If Title=SysAdmin.defAdminUserTitle Or Not SysAdmin.ChangeAdminList Then Response.Write("disabled=""true""") End If%>> 
      </td>
    </tr>
  </table>
</form>
<%End Sub%>
<%Sub AddReco()%>
<form name="form2" method="post" action="?Work=SaveAddReco" onSubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">添加作者信息</td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">作者姓名（笔名）:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title2" size="40" ></td>
    </tr>
    <tr> 
      <td class="BarTitle">性别:</td>
      <td bgcolor="#FFFFFF"><input type="radio" name="Sex" value="0">
        女　 
        <input type="radio" name="Sex" value="1">
        男　 
        <input name="Sex" type="radio" value="2" checked>
        保密</td>
    </tr>
    <tr> 
      <td class="BarTitle">Email:</td>
      <td bgcolor="#FFFFFF"><input name="Email" type="text" class="Input" id="Email" size="40"></td>
    </tr>
    <tr>
      <td class="BarTitle">照片:</td>
      <td bgcolor="#FFFFFF"><input name="BigPhoto" type="text" class="Input" id="BigPhoto" size="40"></td>
    </tr>
    <tr> 
      <td width="25%" valign="top" class="BarTitle">作者简介:</td>
      <td width="75%" bgcolor="#FFFFFF"> <textarea name="Content" cols="60" rows="8" class="Input" id="Content"></textarea></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkAddReco(obj){
    if(obj.Title.value==""){
        alert("请输入[作者姓名（笔名）]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("请输入[作者简介]");
        obj.Content.focus();
        return false;
    }
    return true;    
}
</script> </td>
      <td bgcolor="#FFFFFF"> <input name="Submit4" type="submit" class="button01-out" value="确  定"> 
        <input name="Submit22" type="reset" class="button01-out" value="还  原"> 
        <input name="Submit32" type="button" class="button01-out" value="返  回" onclick="window.history.back();"></td>
    </tr>
  </table>
</form>
<%End Sub%>
</body>
</html>
<%
Sub SaveMdy()
    Dim Sql
        Sql="Select * From AuthorList Where Id=" & Request("Id")
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    If Not (Rs.Eof And Rs.Bof) Then
        Rs("Title")= Trim(Request("Title"))
        Rs("Email")= Trim(Request("Email"))
        Rs("Sex")= Trim(Request("Sex"))
        Rs("Content")= Trim(Request("Content"))
        Rs("BigPhoto")=Trim(Request("BigPhoto"))
        Rs("upTime")= Now
        Rs.Update
    End If
    Rs.Close
    Set Rs=Nothing    
    Response.Redirect("Author_List.asp")
End Sub

Sub DelReco()
    Dim Sql
    Sql="Delete From AuthorList Where Id=" & Request("Id")
    Conn.ExeCute(Sql)
    Response.Redirect("Author_List.asp")
End Sub

Sub SaveAddReco()
    Dim Sql
        Sql="Select Top 1 * From AuthorList Order By ID DESC"
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    Rs.AddNew
    Rs("Title")= Trim(Request("Title"))
    Rs("Email")= Trim(Request("Email"))
    Rs("Sex")= Trim(Request("Sex"))        
    Rs("Content")= Trim(Request("Content"))
    Rs("BigPhoto")=Trim(Request("BigPhoto"))
    Rs("AddTime")= Now
    Rs("upTime")= Now
    Rs.Update
    Rs.Close
    Set Rs=Nothing
    Response.Redirect("Author_List.asp")
End Sub
%>