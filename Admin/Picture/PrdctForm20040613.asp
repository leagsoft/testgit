<%
'定义菜单
cNav = "Prdct"
%>
<!--#include file="Head.asp"-->
<!--#include File="Include/Conn.asp"-->
<!--#include File="Include/Function.asp"-->
<%
'定义按钮
cAdd    = "增加记录"
cSave   = "保存记录"
cDel    = "删除记录"
cList   = "列表显示"

'接收按键
cAction = Trim(Request("cAction"))
cMode   = Trim(Request("cMode"))

'接收变量
nProId      = Request("nProId")
nType1      = Replace(Trim(Request("nType1")),"'","’")
nProName    = Replace(Trim(Request("nProName")),"'","’")
nIntroduct  = Replace(Trim(Request("nIntroduct")),"'","’")
nPic        = Replace(Trim(Request("nPic")),"'","’")

If nProId = "" or cAction=cAdd Then
   cMode = "New"
   cMsg  = "录入新的产品数据！"
Else 
   cMode = "Edit"
End If

'判断按键
If cAction = cSave Then
       '数据判断
       If nType1 = "" Then
          cMsg  = "请选择产品分类！"
       ElseIf nProName = "" Then
	  cMsg  = "产品名称不能为空！"
       ElseIf cMode = "New" Then
          '写入数据库
           cSql="insert into Products(Type1,Proname,Content,Pic,CreDate) values ('"&nType1&"','"&nProName&"','"&nIntroduct&"','"&nPic&"',getdate())"
	  On Error Resume Next
	  Conn.Execute(cSql)
	  If Err.Number = 0 Then
		 cMsg="记录增加成功！"
		 Session("cApply")=""
		 cMode = "Edit"
	  Else
		 cMsg = "数据处理失败！"
		 cMode = "New"
	  End If
	  Err.Clear
       ElseIf cMode = "Edit" Then
          '写入数据库
		cSql="Update Products set Type1='"&nType1&"',Proname='"&nProname&"',Content='"&nIntroduct&"',Pic='"&nPic&"'"&" where PROID="&nProid
		'Response.Write csql
		'Response.End 
	  On Error Resume Next	 
	  Conn.Execute(cSql)
	  If Err.Number = 0 Then
		 cMsg = "记录更改成功！"
		 Session("cApply")=""
	  Else
		 cMsg = "数据处理失败！"
	  End If
	  Err.Clear
       End If  
ElseIf cAction = cDel Then
       cSql = "delete Products where PROID="&nProId&""
       Response.Write csql
       Conn.Execute(cSql)
       If Err.Number = 0 Then
          Response.Redirect "index.asp"
          Session("cApply")=""
       Else
          cMsg = "数据处理失败！"
       End If       
End If

If cMode="New" and cAction <> "" Then
      nProId      = ""
      nType1      = ""
      nProName    = ""
      nIntroduct  = ""
      nPic        = ""
      nAllowSys   = ""
      Session("cApply")=""
ElseIf cMode = "Edit" and nProId<>"" Then
   '查询数据库
   cSql = "select * from Products where PROID = "&nProId
   Set Rs = Conn.Execute(cSql)
   If Rs.Eof Then
      Response.Redirect "Index.asp"
   Else
      '赋予变量
      nType1      = Trim(Rs("Type1"))
      nProName    = Trim(Rs("ProName"))
      nIntroduct  = Trim(Rs("Content"))
      nPic        = Trim(Rs("Pic"))
      'nAllowSys   = Trim(Rs("AllowSys"))
      Session("cApply")=nPic
   End If
   Rs.Close
   Set Rs = Nothing
End If
%>
<script language="JavaScript">
<!--
function UserValidator(theForm)
{
  if (theForm.nType1.value==null)
  {
    alert("请选择图片类别！");
  theForm.nType1.focus();
  return(false);
  }
  if (theForm.nProName.value==null)
  {
    alert("请输入图片标题！");
	theForm.nProName.focus();
	return(false);
  }
  if (theForm.nPic.value==null)
  {
    alert("请上传图片！");
	theForm.nPic.focus();
	return(false);
  }  
return (true);
}
//-->
</script>
<table width="90%" border="1" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFDFDF" bordercolorlight="#333333" bordercolordark="#FFFFFF">
  <form onsubmit="return UserValidator(this)" NAME="PrdctForm" METHOD="Post" ACTION="PrdctForm.asp">
  <input TYPE="Hidden" NAME="nProId" VALUE="<%=nProId%>">
  <input TYPE="Hidden" NAME="cMode" VALUE="<%=cMode%>">
  <tr>
    <td colspan="2" bgcolor="#FFDA99" valign="middle" height="30">&nbsp;<a href="index.asp">图片管理</a> -> <font color=red>图片信息维护</font></td>
  </tr>
  <tr>
    <td colspan="2" height="22" bgcolor="#00DDDD" valign="middle" nowrap>&nbsp;提示信息：<%= cMsg %></td>
  </tr>
  <tr>
    <td height="221" align="center"> 
<table width="95%" border=0 cellpadding=1 cellspacing=0 align="center">
    <tr>
      <td width="20%" align="right"><font color="brown">图片分类：</td>
      <td width="50%">
<!--<select name="nType1" onchange="javascript:document.PrdctForm.submit()">-->
<select name="nType1">
<option value="">请选择类型</option>
<%
cCitySql = "select DICID,Type from SYSDIC order by DICID desc"
Set RsCity = Conn.Execute(cCitySql)
Do
  If RsCity.Eof Then Exit Do
%>
<option value="<%= Trim(RsCity("DICID"))%>"<%If Trim(RsCity("DICID")) = nType1 Then Response.Write " selected" End If%>><%=Trim(RsCity("Type"))%></option>
<%
  RsCity.MoveNext
Loop
RsCity.Close
Set RsCity = Nothing
%>
</select>

      </td>
    </tr>
    <tr>
      <td width="20%" align="right"><font color="brown">图片标题：</td>
      <td width="80%"><input type="text" name="nProName" size="50" maxlength="100" value="<%=nProName%>"></td>
    </tr>
    <tr>
      <td width="20%" align="right"><font color="brown">图片介绍：</td>
      <td width="80%"><textarea rows="10" name="nIntroduct" cols="50" wrap><%=nIntroduct%></textarea></td>
    </tr>
    <tr>
      <td width="20%" align="right"><font color="brown">图片：</td>
      <td width="80%"><INPUT TYPE="Text" Name="nPic" Size="50" value="<%=nPic%>"> <A onClick="window.open('FileSystem/View.asp','ResWin','resizable,scrollbars,width=600,height=500')" style="cursor:hand" title="按此增加图片"> <font color=red><u>增加</u></font> </A></td>
    </tr>
  </table><br> 
  <br>
    </td>
    <td bgcolor="#FFDA99" align="center" VALIGN="Center" height="221">
<%
If cMode = "Edit" Then
%>
      <input TYPE="submit" NAME="cAction" VALUE="<%= cAdd%>"><br><br>
      <input TYPE="submit" NAME="cAction" VALUE="<%= cDel%>"><br><br>
<%
End If
%>
      <input TYPE="submit" NAME="cAction" VALUE="<%= cSave%>"><br><br>
      <input TYPE="Button" NAME="cAction" VALUE="<%= cList %>" onClick="JavaScript:GoToList()">
    </td>
  </tr>
</form>
</table>
</center>
<form NAME="GoList" METHOD="Post" ACTION="Index.asp">
<input TYPE="Hidden" NAME="fnStartRecord" VALUE="<%= fnStartRecord %>">
</form>
<br>
<br>
<!--#include file="Foot.asp"-->


