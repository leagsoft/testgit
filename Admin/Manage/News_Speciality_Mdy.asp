<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!-- #include file="Include/Speciality_Fun.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If

If Not SysAdmin.ChangeSpeciality Then
    Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End If
%>
<html>
<head>
<title>News_Speciality_Mdy.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
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
    Set Rs=Conn.ExeCute("Select * From News_Speciality Where Id=" & Request("Id"))
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("记录未找到")
        Response.End
    End If
    Dim Id,Title,Parent,Explain,IdList
    Id=Rs("Id")
    Title=Rs("Title")
	Parent=Rs("Parent")
    Explain=Rs("Explain")
    Rs.Close
    Set Rs=Nothing
%>
<form name="form1" method="post" action="?Work=SaveMdy" onSubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr align="center"> 
      <td colspan="2" class="BarTitleBg">编辑资源特性</td>
    </tr>
    <tr> 
      <td width="17%" align="right" class="BarTitle">父特性ID:</td>
      <td width="83%" bgcolor="#FFFFFF"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="34%"> <input name="Parent" type="text" class="Input" id="Parent5" value="<%=Parent%>" size="4">
              <input name="Id" type="hidden" id="Parent22" value="<%=Request("Id")%>"></td>
            <td width="66%"><font color="#666666">用于转移[特性]至其它特性下,请慎重更改</font></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">详细位置:</td>
      <td bgcolor="#FFFFFF"> <%=Spec_GetSpecialityPath(Id,"News_Speciality_List.asp")%> </td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">特性名称:</td>
      <td bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title" value="<%=Title%>" size="60"></td>
    </tr>
    <tr> 
      <td align="right" valign="top" class="BarTitle"> <p>特性简介:</p></td>
      <td bgcolor="#FFFFFF"> <textarea name="Explain" cols="60" rows="5" class="Input" id="remark" style="width:100%"><%=Explain%></textarea> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="right"> 
        <script>
function checkMdyReco(obj){
    if(obj.Title.value==""){
        alert("请输入[特性名称]");
        obj.Title.focus();
        return false;
    }
    return true;
}
</script> </td>
      <td><input name="Submit" type="submit" class="button01-out" value="确  定"> 
        <input name="Submit2" type="reset" class="button01-out" value="还  原"> 
        <input name="Submit3" type="button" class="button01-out" value="返  回" onclick="window.history.back();"> 
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="right" bgcolor="#FFFFFF"> <script>
function DelReco(id){
    if(confirm("你确定删除吗？")){
        window.location="?Work=DelReco&Id="+id;
    }
}
</script>
        <input name="Submit5" type="button" class="button01-out" onClick="DelReco(<%=Request("Id")%>)" value="删  除">
      </td>
    </tr>
  </table>
</form>
<%End Sub%>
<%Sub AddReco()%>
<form name="form2" method="post" action="?Work=SaveAddReco" onSubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">添加资源特性</td>
    </tr>
    <tr> 
      <td width="17%" align="right" class="BarTitle">详细位置:</td>
      <td width="83%" bgcolor="#FFFFFF"><input name="Parent" type="hidden" id="Parent" value="<%=Request("Parent")%>">
        <%=Spec_GetSpecialityPath(Request("Parent"),"News_Speciality_List.asp")%></td>
    </tr>
    <tr> 
      <td align="right" class="BarTitle">特性名称:</td>
      <td bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title" size="60"></td>
    </tr>
    <tr> 
      <td align="right" valign="top" class="BarTitle"> <p>特性简介:</p></td>
      <td bgcolor="#FFFFFF"> <textarea name="Explain" cols="60" rows="5" class="Input" id="remark" style="width:100%"></textarea> 
      </td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkAddReco(obj){
    if(obj.Title.value==""){
        alert("请输入[特性名称]");
        obj.Title.focus();
        return false;
    }
}
</script> </td>
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
    Dim Sql
        Sql="Select * From News_Speciality Where Id=" & Request("Id")
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    If Not (Rs.Eof And Rs.Bof) Then
        Rs("Parent")=CLng(Request("Parent"))
        Rs("Title")= Trim(Request("Title"))
        Rs("Explain")= Trim(Request("Explain"))
        Rs("upTime")= Now
        Rs.Update
    End If
    Rs.Close
    Set Rs=Nothing    
    Response.Redirect("News_Speciality_List.asp?Parent="&Request("Parent"))
End Sub

Sub DelReco()
    Dim Sql
    Dim Rs
    Sql="Select Count(*) As Num From News_Speciality Where Parent="&Request("Id")
    Set Rs=Conn.ExeCute(Sql)
    If Rs("Num")>=1 Then
        Response.Write("<script>alert('<操作失败>\n其下还有特性，无法删除！"& SoftCopyright_Script &"');window.history.back();</script>")
        Rs.Close
        Response.End
    Else
        Sql="Delete From News_Speciality Where Id=" & Request("Id")
        Conn.ExeCute(Sql)
    End If
    Response.Redirect("News_Speciality_List.asp")
End Sub

Sub SaveAddReco()
    Dim Sql
        Sql="Select Top 1 * From News_Speciality Order By ID DESC"
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    Rs.AddNew
    Rs("Parent")=CLng(Request("Parent"))
    Rs("Title")= Trim(Request("Title"))
    Rs("Explain")= Trim(Request("Explain"))
    Rs("upTime")= Now
    Rs.Update
    Rs.Close
    Set Rs=Nothing
    Response.Redirect("News_Speciality_List.asp?Parent="&Request("Parent"))
End Sub
%>