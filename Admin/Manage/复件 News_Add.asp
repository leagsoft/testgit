<!--#include file="Include/Conn.asp" -->
<!-- #include file="Include/ClassList_Fun.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#include file="Include/Tkl_StringClass.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="Include/OnlineClass.asp" -->
<!--#Include File="Include/UpdateAdminTime.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class

Call UpdateAdminTime()

Dim cFun
Set cFun=New Tkl_StringClass
QXMC=Session("QXMC")		'取得权限名称
Column=Session("Column")	'用户控制的栏目		'Add Benny

if column="省局" then
	Column="广东监管局"

end if
%>
<html>
<head>
<title>News_Add.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script language="JavaScript" src="Include/Tkl_ClassTree.js" type="text/JavaScript"></script>
<script src="Include/Tkl_Skin.js"></script>
<script src="Library/htmlarea/init_htmlarea.js"></script>
</head>

<body bgcolor="#FFFFFF" leftmargin="5" topmargin="5">
<script src="Include/Tkl_Tooltip.js"></script>
<%
Select Case Request("Work")
    Case "AddReco"
        Call AddReco()
    Case "MdyReco"
        Call MdyReco()
End Select
%>
<%Sub AddReco()%>
<form name="form1" method="post" action="News_Mdy.asp?Work=AddReco" onsubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
    <tr align="center"> 
      <td colspan="2" class="BarTitleBg"> 添加资源 </td>
    </tr>
    <tr> 
      <td width="16%" height="9" valign="top" class="BarTitle">资源类型:</td>
      <td width="84%" height="9" bgcolor="#FFFFFF"> 
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="68%">
    <%	if Session("YHZL")<>"管理员" then  %>
    <script language="javascript">
      var root1
      root1=CreateRoot("myTree1","·请选择[资源类别]")
      <%
	  //Dim CurrentClassIdUsed	  
		  CurrentClassIdUsed=Request.Cookies("ZGW_NewsSys")("CurrentClassIdUsed")
		  
	  If Not IsNumeric(CurrentClassIdUsed) Then
		CurrentClassIdUsed=-1
	  End If
	//	  Call CreateClassTree1(SysAdmin.AdminTopClassId,CLng(CurrentClassIdUsed))
	  Call CreateClassTree1(QXMC,Column,CLng(CurrentClassIdUsed))
	  %>
      </script>
<%	elseif Session("YHZL")="管理员" then%>      
    <script language="javascript">
      var root1
      root1=CreateRoot("myTree1","·请选择[资源类别]")
      <%
	  //Dim CurrentClassIdUsed
		  CurrentClassIdUsed=Request.Cookies("ZGW_NewsSys")("CurrentClassIdUsed")
	  If Not IsNumeric(CurrentClassIdUsed) Then
		CurrentClassIdUsed=-1
	  End If
	  Call CreateClassTree3(SysAdmin.AdminTopClassId,CLng(CurrentClassIdUsed))
	  %>
      </script>
<%end if%> 
</td>
    <td width="32%" align="right" valign="top"><font color="red">*</font>
	<label for="CurrentClassIdUsed" title="下载添加资源时自动使用当前类别">
	<input type="checkbox" id="CurrentClassIdUsed" name="CurrentClassIdUsed" value="1" <%If CurrentClassIdUsed<>-1 Then Response.Write "checked" End If%>>下次使用</label>	</td>
  </tr>
</table></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">资源标题:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="Title" type="text" class="Input" id="Title" size="60">&nbsp;&nbsp;<font color="red">*</font> 
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">跳转链接:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="Url" type="text" class="Input" id="Url" size="60">(不用此功能时请置空) 
      </td>
    </tr>
    <%
    '*******Add By BennyLiu:20040618***********
    '***判断是否是发布局长专题，是的话用下拉框来选择局长。********
    If QXMC="局长专题" then
		cSql="select YHDL,UserRankRight from YHXX where YHBM='局领导' order by UserRankRight"
		set cRs=server.CreateObject ("Adodb.Recordset")
		cRs.Open cSql,connect,1,3    
    %>
    <tr> 
      <td width="16%" class="BarTitle">资源作者:</td>
      <td width="84%" bgcolor="#FFFFFF">
		<select name="Author">
			<%while not cRs.EOF 
				YHMC=Trim(cRs("YHDL"))
				UserRankRight=Trim(cRs("UserRankRight"))
			%>		
			<option value="<%=UserRankRight%>"><%=YHMC%></option>
			<%
				cRs.MoveNext 
			wend
			cRs.Close
			set cRs=nothing
			%>			
		</select>
      &nbsp;&nbsp;<font color="red">*</font></td>
    </tr>    
    <%Else%>
    <tr> 
      <td width="16%" class="BarTitle">资源作者:</td>
      <td width="84%" bgcolor="#FFFFFF"><input type="text" name="Author" class="Input" id="Author" value="<%=session("YHDL")%>">
      &nbsp;&nbsp;<font color="red">*</font></td>
    </tr>
    <%End If%>
    <tr> 
      <td width="16%" class="BarTitle">来源:</td>
      <td width="84%" bgcolor="#FFFFFF"><input type="text" name="From" class="Input" id="From" value="<%=Session("YHBM")%>">&nbsp;&nbsp;<font color="red">*</font></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">责任编辑:</td>
      <td bgcolor="#FFFFFF"><input name="Editor" type="text" class="Input" value="<%=Session("YHDL")%>">
      </td>
    </tr>
<!--Add by BennyLiu:20040712-->    
    <%if QXMC="分局动态" then%>
    <tr> 
      <td width="16%" class="BarTitle">浏览者：</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="添加群组" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=group&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="添加用户" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=user&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="添加群组" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="添加用户" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Session("Browser")%>">
      </td>
    </tr>
    <%end if%> 
<!--End Add-->
    <%'Add By BennyLiu:20040625   为了定义金融统计信息的浏览者
	if QXMC="金融统计信息" then
    %>
    <tr> 
      <td width="16%" class="BarTitle">浏览者：</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="添加群组" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=group&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="添加用户" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=user&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="添加群组" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="添加用户" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Session("Browser")%>">
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">文件类型：</td>
      <td bgcolor="#FFFFFF">
		<select name="DocumentType">
			<option value="0">报表</option>
			<option value="1">文档</option>
		</select>
      </td>
    </tr>
    <%end if%>    
    <tr> 
      <td width="16%" valign="top" class="BarTitle">关键字:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="KeyWord" type="text" class="Input" id="Map" size="50" onmouseover="showToolTip('各[关键词]之间请使用“逗号”隔开，如：<br><b>电脑,游戏</b>',event.srcElement)" onmouseout="hiddenToolTip()"></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">点击:</td>
      <td bgcolor="#FFFFFF"><input name="Count" type="text" class="Input" id="Count" value="1" size="4"></td>
    </tr>
    <tr>
        <td height=23 colspan="2" class="BarTitle">
		<textarea name="logtext" style="display:none" id="body"></textarea>
		<!--#include file="htmedit.asp"-->
		</td>
	</tr>     
    <!--<tr> 
      <td height="22" colspan="2" valign="top" class="BarTitle"><font color="#0000FF">资源内容</font>: 
        <span style="cursor:hand" Title="编辑区加高" onClick="tdNewsContent.style.height=1000;">[加高]</span>&nbsp;<span style="cursor:hand" Title="编辑区默认高度" onClick="tdNewsContent.style.height=400;">[默认]</span>&nbsp;<span style="cursor:hand" Title="编辑区减低" onClick="tdNewsContent.style.height=200;">[减低]</span>&nbsp;<span style="cursor:hand" Id="ShowHiddenHtmlEdit" Title="显示/隐藏" onClick="if(trNewsContent.style.display==''){trNewsContent.style.display='none';ShowHiddenHtmlEdit.innerHTML='[显示]'}else{trNewsContent.style.display='';ShowHiddenHtmlEdit.innerHTML='[隐藏]'}">[隐藏] 
        </span>&nbsp;&nbsp;<font color="red">*</font></td>
    </tr>
    <tr id="trNewsContent"> 
      <td height="400" colspan="2" valign="top" bgcolor="buttonface" Id="tdNewsContent1"><textarea name="NewsContent"></textarea></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">上传:</td>
      <td bgcolor="#FFFFFF"><input name="SmallImg" type="text" class="Input" id="SmallImg" value="" size="50">
        <input type="button" name="Button" value="..." onClick="window.open('FileSystem/View.asp','ResWin','resizable,scrollbars,width=600,height=500')"></td>
    </tr>-->
    <tr bgcolor="#FFFFFF"> 
      <td height="13"> 
        <script>
function checkAddReco(obj){
    var IsNewsClassChecked=false
    var Coll
        Coll=obj.item("radioBoxItem")
	if(Coll.length)
	{
		for(var i=0;i<Coll.length;i++)
		{
			if(Coll.item(i).name=="radioBoxItem"){
				if(Coll.item(i).checked){
					IsNewsClassChecked=true
					break
				}
			}
		}
	}else{
		IsNewsClassChecked=Coll.checked
	}
    if(!IsNewsClassChecked){
        alert("请选择[资源类别]")
        return false
    }
    if(obj.Title.value==""){
        alert("请输入[资源标题]");
        obj.Title.focus();
        return false;
    }
    //if(obj.Url.value!=""){
        //if(obj.Url.value.search(/^[a-z0-9]+:\/\/[a-z0-9]+/i)==-1)
       // {
       //     alert("[跳转链接]格式有误");
      //      obj.Url.focus();
     //       return false;
    //    }
   // }else{
    
    //}
   // if(obj.Author.value==""){
   //     alert("请选择[资源作者]");
   //     obj.Author.focus();
   //     return false;
   // }
    //if(obj.From.value==""){
     //   alert("请选择[资源来源]");
    //    obj.From.focus();
    //    return false;
    //}
    //if(obj.KeyWord.value==""){
       // alert("请输入[资源关键字]");
        //obj.KeyWord.focus();
        //return false;
    //}        
    if(obj.Count.value==""){
        alert("请输入[资源点击率]");
        obj.Count.focus();
        return false;
    }
    if(obj.Url.value=="")
    {
        if(obj.logtext.value==""){
            alert("请输入[资源内容]");
            return false;
        }
    }
    //if(obj.ShortContent.value==""){
    //    alert("请输入[资源摘要]");
	//	obj.ShortContent.focus();
    //    return false;
    //}
    obj.SaveAddButton.disabled=true
    return true;
}
</script> </td>
      <td><input name="SaveAddButton" type="submit" id="SaveAddButton" class="button01-out" value="确  定"> 
        <input name="Submit2" type="reset" class="button01-out" value="还  原"> 
        <input name="Submit3" type="button" class="button01-out" value="返  回" onclick="window.history.back();"></td>
    </tr>
  </table>
</form>
<script language="javascript1.2">
 
//editor_generate('NewsContent',config);
</script>
<%End Sub%>
<%
Sub MdyReco()
    Dim Sql
        Sql="Select * From News Where Id="&Request("Id")
		Session("CurrentEdit_ResourceId")=Request("Id")
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    'Response.Write Rs("Author")
%>
<form name="form2" method="post" action="News_Mdy.asp?Work=SaveMdy" onsubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr align="center"> 
      <td colspan="2" class="BarTitleBg"> 编辑资源 
        <input name="Id" type="hidden" id="Id" value="<%=Rs("Id")%>"></td>
    </tr>
    <tr> 
      <td width="16%" height="9" valign="top" class="BarTitle">资源类型:</td>
      <td width="84%" height="9" bgcolor="#FFFFFF">
<%if Session("YHZL")<>"管理员" then%>      
        <script>
      var root2
      root2=CreateRoot("myTree2","·请选择[资源类别]")
      <%
      //Call CreateClassTree2(SysAdmin.AdminTopClassId,Rs("Class"))
      Call CreateClassTree2(QXMC,Column,Rs("Class"))
      %>
      </script>
<%elseif Session("YHZL")="管理员" then%>
        <script>
      var root2
      root2=CreateRoot("myTree2","·请选择[资源类别]")
      <%
      Call CreateClassTree4(SysAdmin.AdminTopClassId,Rs("Class"))
      %>
      </script>
<%end if%>
      <font color="red">*</font> 
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">资源标题:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="Title" type="text" class="Input" id="Title" size="60" value="<%=cFun.HTMLEncode2(Rs("Title"))%>">&nbsp;&nbsp;<font color="red">*</font>  
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">跳转链接:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="Url" type="text" class="Input" id="Url" size="60" value="<%=cFun.HTMLEncode2(Rs("Url"))%>">(不用此功能时请置空) 
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">资源作者:</td>
      <%
      'Modify By BennyLiu:20040625
      If QXMC="局长专题" then
			cSql2="select YHDL,UserRankRight from YHXX where YHBM='局领导' order by UserRankRight"
			set cRs2=server.CreateObject ("Adodb.Recordset")
			cRs2.Open cSql2,connect,1,3
	  %>
	  <td width="84%" bgcolor="#FFFFFF">
				<select name="Author">
	  <%
			while not cRs2.EOF 
			YHMC=Trim(cRs2("YHDL"))
			UserRankRight=Trim(cRs2("UserRankRight"))
      %>
			
				<option value="<%=UserRankRight%>"<%if UserRankRight=Trim(Rs("Author")) then Response.Write " selected" end if%>><%=YHMC%></option>
	  <%
				cRs2.MoveNext 
			wend
			cRs2.Close
			set cRs2=nothing
	  %>
			</select>&nbsp;&nbsp;<font color="red">*</font></td>
      <%Else%>	        
      <td width="84%" bgcolor="#FFFFFF"><input type="text" name="Author" class="Input" id="Author" value="<%=session("YHDL")%>">&nbsp;&nbsp;<font color="red">*</font></td>
      <%End If%>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">来源:</td>
      <td width="84%" bgcolor="#FFFFFF"><input type="text" name="From" class="Input" id="From" value="<%=session("YHBM")%>">&nbsp;&nbsp;<font color="red">*</font></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">责任编辑:</td>
      <td bgcolor="#FFFFFF"><input name="Editor" type="text" class="Input" value="<%=Session("YHDL")%>">
      </td>
    </tr>
<!--Add by BennyLiu:20040712-->
    <%if QXMC="分局动态" then%>
    <tr> 
      <td width="16%" class="BarTitle">浏览者:</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="添加群组" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=group&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="添加用户" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=user&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="添加群组" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="添加用户" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Rs("Browser")%>">
      </td>
    </tr>   
    <%end if%>
<!--End Add-->    
    <%'Add by BennyLiu:20040625  为了定义可浏览者，只有金融统计信息才能此项。
	if QXMC="金融统计信息" then
    %>
    <tr> 
      <td width="16%" class="BarTitle">浏览者:</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="添加群组" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=group&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="添加用户" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=user&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="添加群组" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="添加用户" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Rs("Browser")%>">
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">文件类型：</td>
      <td bgcolor="#FFFFFF">
		<select name="DocumentType">
			<option value="0"<%if Rs("IsDocument")="0" then Response.Write " selected" end if%>>报表</option>
			<option value="1"<%if Rs("IsDocument")="1" then Response.Write " selected" end if%>>文档</option>
		</select>
      </td>
    </tr>    
    <%end if%>    
    <tr> 
      <td width="16%" valign="top" class="BarTitle">关键字:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="KeyWord" type="text" class="Input" id="Map" size="50" value="<%=cFun.HTMLEncode2(Rs("KeyWord"))%>">
        (各关键词间用[逗号]隔开) </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">点击:</td>
      <td bgcolor="#FFFFFF"><input name="Count" type="text" class="Input" id="Count" size="4" value="<%=Rs("Count")%>"></td>
    </tr>
    <tr>
        <td height=23 colspan="2" class="BarTitle">
		<textarea name="logtext" style="display:none" id="body">
		<%If (Not IsNull(Rs("Content"))) Or (Not ""<>Rs("Content")) Then 
			Response.Write(Server.HtmlEnCode(replace(Rs("Content"),"11.36.19.2","10.100.0.2")))
		end If%>
		</textarea>
		<!--#include file="htmedit.asp"-->
		</td>
	</tr>     
    <!--<tr> 
      <td height="19" colspan="2" valign="top" class="BarTitle"><font color="#0000FF">资源内容</font>: 
        <span style="cursor:hand" Title="编辑区加高" onClick="tdNewsContent.style.height=1000;">[加高]</span>&nbsp;
        <span style="cursor:hand" Title="编辑区默认高度" onClick="tdNewsContent.style.height=400;">[默认]</span>&nbsp;
        <span style="cursor:hand" Title="编辑区减低" onClick="tdNewsContent.style.height=200;">[减低]</span>&nbsp;
        <span style="cursor:hand" Id="ShowHiddenHtmlEdit" Title="显示/隐藏" onClick="if(trNewsContent.style.display==''){trNewsContent.style.display='none';ShowHiddenHtmlEdit.innerHTML='[显示]'}else{trNewsContent.style.display='';ShowHiddenHtmlEdit.innerHTML='[隐藏]'}">[隐藏]</span>
		&nbsp;&nbsp;<font color="red">*</font> 
      </td>
    </tr>
    <tr id="trNewsContent"> 
      <td height="400" colspan="2" valign="top" bgcolor="buttonface" Id="tdNewsContent"> <textarea name="NewsContent"><%If (Not IsNull(Rs("Content"))) Or (Not ""<>Rs("Content")) Then Response.Write Server.HtmlEnCode(replace(Rs("Content"),"11.36.19.2","10.100.0.2"))End If%></textarea></td>
    </tr>
    <tr> 
      <td class="BarTitle">上传:</td>
      <td bgcolor="#FFFFFF"><input name="SmallImg" type="text" class="Input" id="SmallImg" value="<%=cFun.HTMLEncode2(Rs("SmallImg"))%>" size="50">
        <input type="button" name="Button3" value="..." onClick="window.open('FileSystem/View.asp','ResWin','resizable,scrollbars,width=600,height=500')"></td>
    </tr>-->
    <tr bgcolor="#FFFFFF"> 
      <td height="27"> 
        <script>
function checkMdyReco(obj){
    var IsNewsClassChecked=false
    var Coll
        Coll=obj.item("radioBoxItem")
	if(Coll.length)
	{
		for(var i=0;i<Coll.length;i++)
		{
			if(Coll.item(i).name=="radioBoxItem"){
				if(Coll.item(i).checked){
					IsNewsClassChecked=true
					break
				}
			}
		}
	}else{
		IsNewsClassChecked=Coll.checked
	}
    if(!IsNewsClassChecked){
        alert("请选择[资源类别]")
        return false
    }
    if(obj.Title.value==""){
        alert("请输入[资源标题]");
        obj.Title.focus();
        return false;
    }
    //if(obj.Url.value!=""){
       // if(obj.Url.value.search(/^[a-z0-9]+:\/\/[a-z0-9]+/i)==-1)
       // {
       //     alert("[跳转链接]格式有误");
      //      obj.Url.focus();
     //       return false;
     //   }
    //}
    //if(obj.Author.value==""){
      //  alert("请选择[资源作者]");
      //  obj.Author.focus();
     //   return false;
    //}
    //if(obj.From.value==""){
       // alert("请选择[资源来源]");
      //  obj.From.focus();
     //   return false;
    //}
    //if(obj.KeyWord.value==""){
      //  alert("请输入[资源关键字]");
       // obj.KeyWord.focus();
       // return false;
   // }        
    if(obj.Count.value==""){
        alert("请输入[资源点击率]");
        obj.Count.focus();
        return false;
    }    
    if(obj.Url.value=="")
    {
        if(obj.logtext.value==""){
            alert("请输入[资源内容]");
            return false;
        }
    }
    //if(obj.ShortContent.value==""){
    //    alert("请输入[资源摘要]");
	//	obj.ShortContent.focus();
    //} 
    form2.SaveMdyButton.disabled=true
    return true;
}
</script></td>
      <td><input name="Submit4" type="submit" id="SaveMdyButton" class="button01-out" value="确  定"> 
        <input name="Submit22" type="reset" class="button01-out" value="还  原"> 
        <input name="Submit32" type="button" class="button01-out" value="返  回" onclick="window.history.back();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="27">&nbsp;</td>
      <td align="right"> <script>
function DelReco(id,rDel,info){
    if(confirm(info)){
            window.location="News_Mdy.asp?Work=DelReco&Id="+id+"&RealDel="+rDel;
    }
}
</script> <%If Not CBool(Rs("Del")) Then%> 
        <table width="26%" border="0" cellspacing="1" cellpadding="0">
          <tr> 
            <td align="right"><label for="RealDel"><input name="RealDel" type="checkbox" id="RealDel" value="1"> 
              <font color="#0000FF">彻底删除</font></label> <input name="Submit322" type="button" class="button01-out" value="删  除" onclick="DelReco('<%=Rs("Id")%>',form2.RealDel.checked,'你确定删除吗？')"></td>
          </tr>
        </table>
        <%Else%> 
        <table width="26%" border="0" cellspacing="1" cellpadding="0">
          <tr> 
            <td align="right"><input name="Submit3223" type="button" class="button01-out" value="恢  复" onclick="DelReco('<%=Rs("Id")%>','0','你确定[恢复]此记录吗？')" Title="[恢复]当前已经被[虚拟删除]的记录"> 
              <input name="Submit3222" type="button" class="button02-out" value="彻底删除" onclick="DelReco('<%=Rs("Id")%>','1','你确定[彻底删除]吗？\n[彻底删除]的记录将无法被[恢复]')" title="[彻底删除]的记录将无法还原"></td>
          </tr>
        </table>
        <%End If%> </td>
    </tr>
  </table>
</form>
<script language="javascript1.2">

//editor_generate('NewsContent',config);
</script>
<%
    Rs.Close
End Sub
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display='';window.scrollTo(window.pageXOffset,2000);}else{HelpTab.style.display='none'}">&nbsp;<img src="Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:none"> 
    <td width="2%">&nbsp;</td>
    <td width="98%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <!--<tr> 
          <td>·生成/更新静态文件:</td>
        </tr>
        <tr> 
          <td> 　　未审核的资源不能执行此操作</td>
        </tr>-->
        <tr> 
          <td>·'删除'、'恢复'、'彻底删除'</td>
        </tr>
        <tr> 
          <td> 　　1.'删除':保将记录在逻辑标记为'已删除',可通过'回收站'中恢复该记录。'删除'的同时系统也将相对应的静态资源文件进行删除</td>
        </tr>
        <tr> 
          <td>　　2.'彻底删除':将无法挽回被'彻底删除'的资源记录</td>
        </tr>
        <tr> 
          <td>·自动分页标签:</td>
        </tr>
        <tr> 
          <td>　　将该标签放入相应的Html源码位置当中,系统就以该分页标签所在位置将资源内容分割成多个页面进行生成</td>
        </tr>
        <tr> 
          <td>·跳转链接:</td>
        </tr>
        <tr> 
          <td>　　允许资源在被点击后直接跳转到指定的[跳转链接]地址，若不使用此功能请置空。</td>
        </tr>
        <tr> 
          <td>·资源标题样式:</td>
        </tr>
        <tr> 
          <td>　　资源标题允许使用Html标签进行色彩、风格的设定，如：“&lt;font color='red'&gt;××××新闻标题&lt;/font&gt;”，但注意，所有的双引号必须改成单引号。</td>
        </tr>
      </table>
      <a name="Help"></a></td>
  </tr>
</table>
</body>
</html>
<%
Sub SpecialityList_Add(ParentId)
    Dim Sql
        Sql="Select * From News_Speciality Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        If Rs("Parent")=0 Then
            Response.Write "root3.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""checkbox\"" NAME=\""Speciality\"" value=\"""&Rs("Id")&"\""><span onmouseover=\""showToolTip('"&Replace(Rs("Explain"),"""","'")&"',event.srcElement)\"" onmouseout=\""hiddenToolTip()\"">"&Rs("Title")&"</span>"")" & vbCrLf
        Else
            Response.Write "root3.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""checkbox\"" NAME=\""Speciality\"" value=\"""&Rs("Id")&"\""><span onmouseover=\""showToolTip('"&Replace(Rs("Explain"),"""","'")&"',event.srcElement)\"" onmouseout=\""hiddenToolTip()\"">"&Rs("Title")&"</span>"")" & vbCrLf
        End If
        Call SpecialityList_Add(Rs("Id"))
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

Sub SpecialityList_Mdy(ParentId,itemList)
    Dim Sql
        Sql="Select * From News_Speciality Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    Dim radioSelected
    While Not Rs.Eof
        If Instr(","&itemList&",",","&Rs("Id")&",")<>0 Then
            radioSelected="checked"
        Else
            radioSelected=""
        End If
        If Rs("Parent")=0 Then
            Response.Write "root4.CreateNode("&Rs("Id")&",-1,""<INPUT "&radioSelected&" TYPE=\""checkbox\"" NAME=\""Speciality\"" value=\"""&Rs("Id")&"\""><span onmouseover=\""showToolTip('"&Replace(Rs("Explain"),"""","'")&"',event.srcElement)\"" onmouseout=\""hiddenToolTip()\"">"&Rs("Title")&"</span>"")" & vbCrLf
        Else
            Response.Write "root4.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT "&radioSelected&" TYPE=\""checkbox\"" NAME=\""Speciality\"" value=\"""&Rs("Id")&"\""><span onmouseover=\""showToolTip('"&Replace(Rs("Explain"),"""","'")&"',event.srcElement)\"" onmouseout=\""hiddenToolTip()\"">"&Rs("Title")&"</span>"")" & vbCrLf
        End If
        Call SpecialityList_Mdy(Rs("Id"),itemList)
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

'生成栏目树
Sub CreateClassTree1(QXMC,Column,CuClassId)
	Dim radioSelected
	Dim Rs
    Dim Sql

	if Column="" then
		Sql="Select * from Classlist where Title='"&QXMC&"'"
	end if		
	if Column<>"" then
		Sql="Select * from Classlist where Title='"&QXMC&"'"
		set Rs=conn.ExeCute(sql)
		AClassId=Rs("ID")
		Rs.close
		sql="Select * from ClassList where Parent='"&AClassId&"' and Title='"&Column&"'"
	end if
'        Sql="Select * From ClassList Where Title='"&Column&"'"
		set Rs=Conn.ExeCute(sql)
		cClassId=Rs("ID")
		Rs.close
		sql="select * from ClassList where Parent="&cClassId
    Set Rs=Conn.ExeCute(Sql)
'Add benny
	If Rs.eof or Rs.bof then 'Add benny
		sql="select * from classlist where Id="&cClassId
		Dim cRs
		set cRs=Conn.Execute(sql)
		
		while not cRs.eof 
        If CuClassId=cRs("Id") Then
            radioSelected="checked"
        Else
            radioSelected=""
        End If
        Response.Write "    root1.CreateNode("&cRs("Id")&",-1,""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&cRs("Id")&"\"">"&cRs("Title")&""")" & vbCrLf
        cRs.movenext
		wend
		cRs.close
		set cRs=nothing
    Else		'Add benny
'End Add    
'**    Dim radioSelected
    While Not Rs.Eof
        If CuClassId=Rs("Id") Then
            radioSelected="checked"
        Else
            radioSelected=""
        End If
'*        If Rs("Parent")=SysAdmin.AdminTopClassId Then
        If Rs("Parent")=cClassId Then
            Response.Write "    root1.CreateNode("&Rs("Id")&",-1,""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "    root1.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
     '**   CreateClassTree1 Rs("Id"),CuClassId	'Del benny
        Rs.MoveNext
    Wend
	End If		'Add benny    
    Rs.Close
    Set Rs=Nothing
End Sub

'生成栏目树
Sub CreateClassTree2(QXMC,Column,ClassId)
    Dim radioSelected
    Dim Rs
    Dim Sql
	if Column="" then
		Sql="Select * from Classlist where Title='"&QXMC&"'"
	end if		
	if Column<>"" then
		Sql="Select * from Classlist where Title='"&QXMC&"'"
		set Rs=conn.ExeCute(sql)
		AClassId=Rs("ID")
		Rs.close
		sql="Select * from ClassList where Parent='"&AClassId&"' and Title='"&Column&"'"
	end if
'        Sql="Select * From ClassList Where Title='"&Column&"'"
		set Rs=Conn.ExeCute(sql)
		cClassId=Rs("ID")
		Rs.close
		sql="select * from ClassList where Parent="&cClassId
    Set Rs=Conn.ExeCute(Sql)
'Add benny
	If Rs.eof or Rs.bof then 'Add benny
		sql="select * from classlist where Id="&cClassId
		Dim cRs
		set cRs=Conn.Execute(sql)
		
		while not cRs.eof 
        If ClassId=cRs("Id") Then
            radioSelected="checked"
        Else
            radioSelected=""
        End If
        Response.Write "    root2.CreateNode("&cRs("Id")&",-1,""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&cRs("Id")&"\"">"&cRs("Title")&""")" & vbCrLf
        cRs.movenext
		wend
		cRs.close
		set cRs=nothing
    Else		'Add benny
'End Add     
    '**Dim radioSelected
    While Not Rs.Eof
        If ClassId=Rs("Id") Then
            radioSelected="checked"
        Else
            radioSelected=""
        End If
        If Rs("Parent")=cClassId Then
            Response.Write "root2.CreateNode("&Rs("Id")&",-1,""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root2.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        '**CreateClassTree2 Rs("Id"),ClassId	'Del benny
        Rs.MoveNext
    Wend
    End If		'Add benny
    Rs.Close
    Set Rs=Nothing
End Sub

'生成栏目树
Sub CreateClassTree3(ParentId,CuClassId)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    Dim radioSelected
    While Not Rs.Eof
        If CuClassId=Rs("Id") Then
            radioSelected="checked"
        Else
            radioSelected=""
        End If
        If Rs("Parent")=SysAdmin.AdminTopClassId Then
            Response.Write "    root1.CreateNode("&Rs("Id")&",-1,""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "    root1.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        CreateClassTree3 Rs("Id"),CuClassId
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

'生成栏目树
Sub CreateClassTree4(ParentId,ClassId)
    Dim Sql
        Sql="Select Id,Parent,Title From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    Dim radioSelected
    While Not Rs.Eof
        If ClassId=Rs("Id") Then
            radioSelected="checked"
        Else
            radioSelected=""
        End If
        If Rs("Parent")=SysAdmin.AdminTopClassId Then
            Response.Write "root2.CreateNode("&Rs("Id")&",-1,""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root2.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT "&radioSelected&" TYPE=\""radio\"" NAME=\""radioBoxItem\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        CreateClassTree4 Rs("Id"),ClassId
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub
%>