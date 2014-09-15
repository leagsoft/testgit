<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/ClassList_Fun.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="Include/OnlineClass.asp" -->
<!--#Include File="Include/UpdateAdminTime.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not CBool(SysAdmin.Logined) Then
'    Response.Redirect("Login.asp")
'End If

Call UpdateAdminTime()
QXMC=Session("QXMC")	'类别名称
YHZL=Session("YHZL")	'用户类别
UserName=Session("YHDL")	'用户名称

Purview = Session("Purview")	'用户所属权限角色

Column = Session("Column")		'第二级类别名称
if Column="省局" then Column="广东监管局" end if
if Column<>"" then
	cClassID=GetClassID(QXMC)	
end if
%>
<html>
<head>
<title>News_List.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--
function selectAllCheckBox(obj)
{
	if(obj.length)
	{
		for(var i=0;i<obj.length;i++)
		{
			obj[i].checked=!obj[i].checked
		}
	}else{
		
		obj.checked=!obj.checked
	}
}
function chkCheckBox(obj)
{
    var result=0
    if(obj.length)
    {
        for(var i=0;i<obj.length;i++)
        {
            if(obj[i].checked)
            {
                result++
            }
        }
    }else{
        result=obj.checked?1:0
    }
    return result
}

//函数：删除资源
//参数：表单对像,是否真实删除
function DeleteReco(obj,eventObj,mRealDel)
{
	var selNum
		selNum=chkCheckBox(obj.Id)
	if(selNum==0)
	{
		alert("请选择你要[删除/彻底删除/救回]的资源")
		return false
	}
	if(confirm("你确定要[删除/彻底删除/救回]选中的（"+selNum+"）条资源？"))
	{
		obj.Work.value="DelReco"
		obj.RealDel.value=mRealDel
		eventObj.disabled=true
		obj.submit()
	}
}
//函数：审核资源（暂时不使用此功能）
//参数：表单对像
function CheckReco(obj,eventObj)
{
    var selNum
        selNum=chkCheckBox(obj.Id)
    if(selNum==0)
    {
        alert("请选择你要[审核]的资源")
        return false
    }
    obj.Work.value="CheckReco"
    eventObj.disabled=true
    obj.submit()
}

//函数：生成资源（暂时不使用此功能）
//参数：表单对像
function CreateFile(obj,eventObj)
{
    var selNum
        selNum=chkCheckBox(obj.Id)
    if(selNum==0)
    {
        alert("请选择你要[生成]的资源")
        return false
    }
    if(obj.Id.length)
    {
        for(var i=0;i<obj.Id.length;i++)
        {
            if(obj.Id[i].checked && obj.Id[i].HaveChecked!="True")
            {
                alert("当前所选新闻未完全[审核]通过")
                obj.Id[i].focus()
                return false
            }
        }
    }else{
        if(obj.Id.checked && obj.Id.HaveChecked!="True")
        {
            alert("当前所选新闻未完全[审核]通过")
            obj.Id.focus()
            return false
        }
    }
    obj.Work.value="CreateSelectedFile"
    eventObj.disabled=true
    obj.submit()
}

//函数：编辑资源
//参数：Button,记录Id
function MdyReco(Id)
{
    event.srcElement.disabled=true
    document.body.innerHTML="<div align='center'>请稍等,系统正在从数据库中读取您所要编辑的资源信息...</div>"
    window.location="News_Add.asp?Work=MdyReco&Id="+Id
}

function chkSearchForm(obj)
{
    obj.SearchButton.disabled=true
    return true
}
//-->
</SCRIPT>
</head>
<body bgcolor="#FFFFFF">
<%
Dim Parent
If Request("Parent")="" Then
    Parent=SysAdmin.AdminTopClassId
Else
    Parent=CLng(Request("Parent"))
End If
Dim Work
    Work=Request("Work")
Response.Cookies("ZGW_NewsSys")("News_List_WorkType")=Work
Dim sType
    sType=Replace(Request("sType"),"'","")
    If sType="" Then
        sType="Title"
    End If
Dim sKey
    sKey=Replace(Request("sKey"),"'","")
Call ClassList()
Sub ClassList()
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellpadding="3" cellspacing="0" class="BarText">
        <tr> 
          <td width="81%" bgcolor="#f6f6f6">当前位置：<%=GetClassPath2(SysAdmin.AdminTopClassId,Parent,"")%></td>
          <td width="19%" align="right" bgcolor="#f6f6f6"><a href="#AdvanceSh">[搜索] 
            &nbsp; </a> </td>
        </tr>
        <tr> 
          <td colspan="2"> 
<%
Dim Sql,Rs
IF YHZL="管理员" Then
            Sql="Select * From ClassList Where Parent="&Parent&" Order By UpTime"
Else		
'****只有一层类别*******
		If Column="" then
			set Rs1=server.CreateObject ("Adodb.Recordset")
			Sql2="select * from ClassList where Title='"&QXMC&"'"
			
		
		
			Rs1.Open Sql2,conn,1,3
			if not rs1.eof then
			Parent=Rs1("ID")		
			Rs1.Close 
			set Rs1=nothing		
				end if
'****有几层类别*********
		Else
			if Parent<cClassID then
				Parent=cClassID
			end if
		End If
'****管理者、编辑者的查询语句******		
			Sql="select * from classlist where Parent="&Parent
End IF            
            Set Rs=Conn.ExeCute(Sql)
            If Rs.Eof And Rs.Bof Then
                Response.Write("<font color='#666666'>无子类别</font>")
            End If
            While Not Rs.Eof
                Response.Write("<a href='?Parent="&Rs("Id")&"'>"&Rs("Title")&"</a>"&" | ")
                Rs.MoveNext
            Wend
            Rs.Close
%>
          </td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr> 
    <td align="center" class="BarText"> 
<%
        Select Case Work
            Case "Dustbin"
                'Response.Write "<font color=""#FF0000"">资源回收站列表</font>&nbsp;&nbsp;[<A HREF=""#"" onclick=""if(confirm('是否确定要清空[回收站]')){window.location='News_Mdy.asp?Work=ClearDustbin'}else{return false}""><font color=""#33CC00"">清空回收站</font></A>]"
                Response.Write "<font color=""#FF0000"">资源回收站列表</font>"
            Case "UnChecked"
                Response.Write "<font color=""#FF0000"">未审核资源列表</font>"      
            Case Else
                Response.Write "<font color=""#FF0000"">常规资源列表</font>"
        End Select
%>
    </td>
  </tr>
</table>
<%End Sub%>
<%
Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
Dim Sql
    '*Sql=ExeSql()
   
    Sql=ExeSql(YHZL)
  
    Rs.PageSize=Sys_PageSize
    Rs.CacheSize=Rs.PageSize
        
    Rs.Open Sql,Conn,1,1

 
Dim CurrentPage
    If Request("CurrentPage")="" Then
        CurrentPage=1
    Else
        CurrentPage=Request("CurrentPage")
    End If
    Response.Cookies("ZGW_NewsSys")("News_List_CurrentPage")=CurrentPage
    If Not(Rs.Eof And Rs.Bof) Then
        Rs.AbsolutePage=CurrentPage
    End If
%>
<form name="form2" method="post" action="News_Mdy.asp">
  <table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg" style="table-layout:fixed; word-break :break-all;width:100%">
  <tr class="BarTitleBg"> 
      <td width="5%" height="24" align="center">ID</td>
      <td width="50%">资源标题</td>
    <td width="24%" align="center">所属主类别</td>
    <td width="11%" align="center">添加时间</td>
    <td width="10%" align="center">编辑</td>
  </tr>
<%
  Dim I
  For I=1 To Rs.PageSize
      If Rs.Eof Then
        Exit For
    End If
%>
  <tr> 
      <td width="5%" height="24" class="BarTitle" align="center"> 
        <input type="checkbox" name="Id" value="<%=Rs("Id")%>" title="记录编号：<%=Rs("Id")%>" HaveChecked="<%=CBool(Rs("IsChecked"))%>">
      </td>
      <td width="50%" bgcolor="#FFFFFF" title="展开操作面板" onclick="ShowControlPane(window.trNews_<%=Rs("Id")%>)"><label title="责任编辑:<%=Rs("EditorTitle")%>"><%=Rs("title")%></label></td>
    <td width="24%" align="center" bgcolor="#FFFFFF"><%=Rs("ClassTitle")%></td>
    <td width="11%"  align="center" bgcolor="#FFFFFF" title="最后更新:<%=FormatDateTime(Rs("UpTime"),1)%>"><%=FormatDateTime(Rs("AddTime"),2)%></td>
    <td width="10%" align="center" bgcolor="#FFFFFF">
        <input name="Submit2" type="button" class="button01-out" value="编 辑" onClick="MdyReco('<%=Rs("Id")%>')">
      </td>
  </tr>
  <tr Id="trNews_<%=Rs("Id")%>" <%If Not Def_ShowNewsContorlPlane Then Response.Write "style=""display:none""" End If%>>
  <td colspan="5" bgcolor="#FFFFFF">
<%
  If CBool(Rs("Del")) Then
    Response.Write "&nbsp;<a href=""News_Mdy.asp?Work=DelReco&Id="&Rs("Id")&""" title=""恢复资源"">恢复资源</a>"
    Response.Write "&nbsp;<a href=""News_Mdy.asp?Work=DelReco&Id="&Rs("Id")&"&RealDel=1"" title=""不可恢复"">彻底删除</a>"
  Else
    Response.Write "&nbsp;<a href=""News_Mdy.asp?Work=DelReco&Id="&Rs("Id")&""" title=""放入回收站"">删除资源</a>"
  End If
   'Response.Write "&nbsp;<a href=""Comment_List.asp?Work=ByNews&sType=ResId&sKey="&Rs("Id")&""">查看评论</a>"
%>
  </td>
  </tr>
<%
      Rs.MoveNext
  Next
%>
<%If Rs.Eof And Rs.Bof Then%>
  <tr>
  <td align="center" colspan="7" bgcolor="#f6f6f6">暂无相关记录</td>
  </tr>
<%End If%>
</table>
  <table width="100%" border="0" cellpadding="5" cellspacing="0">
    <tr> 
      <td align="right">
        <script src="Include/Tkl_PageList.js"></script>
        <script>Tkl_PageListBar(<%=Rs.PageCount%>,<%=CurrentPage%>,"Work=<%=Work%>&sType=<%=sType%>&sKey=<%=sKey%>&Parent=<%=Parent%>")</script>
      </td>
    </tr>
    <tr>
      <td align="right"> 
        <input type="hidden" name="Work">
        <input type="hidden" name="RealDel">
        <label for="selectAllReco"> 
        <input type="checkbox" name="checkbox2" value="checkbox" id="selectAllReco" onclick="selectAllCheckBox(form2.Id)">
        反选</label> 
        <%If Work="Dustbin" Then%>
        <input name="Submit3223" type="button" class="button01-out" value="恢  复" onClick="DeleteReco(form2,event.srcElement,0)" title="[救回]当前已经被[虚拟删除]的记录">
        <input name="Submit3222" type="button" class="button02-out" value="彻底删除" onClick="DeleteReco(form2,event.srcElement,1)" title="[彻底删除]的资源将无法还原">
        <%Else%>
        <input name="Submit222" type="button" class="button01-out" value="删 除" title="批量删除至回收站" onClick="DeleteReco(form2,event.srcElement,0)">
        <%End If%>
      </td>
    </tr>
  </table>
</form>
<%
Rs.Close
Set Rs=Nothing
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <form name="form1" method="post" action="?" onsubmit="return chkSearchForm(this)">
    <tr bgcolor="#FFFFFF"> 
      <td width="67%" align="right"><a name="AdvanceSh"></a> 
        <input name="Work" type="hidden" id="Work" value="<%=Work%>">
        搜索: 
        <select name="sType" class="Input">
          <option value="Title" <%If sType="Title" Then Response.Write("selected") End If%>>资源标题</option>
          <option value="ClassTitle" <%If sType="ClassTitle" Then Response.Write("selected") End If%>>资源类型</option>
          <option value="Content" <%If sType="Content" Then Response.Write("selected") End If%>>资源内容</option>
          <option value="AuthorTitle" <%If sType="AuthorTitle" Then Response.Write("selected") End If%>>作者名称</option>
          <option value="FromTitle" <%If sType="FromTitle" Then Response.Write("selected") End If%>>资源来源</option>          
        </select> <input name="Parent" type="radio" class="Input" value="0" checked>
        所有 
        <input name="Parent" type="radio" class="Input" value="<%=Parent%>">
        当前 </td>
      <td width="25%" align="right"> <input name="sKey" type="text" class="Input" id="sKey" style="width:100%" value="<%=Trim(Request("sKey"))%>"></td>
      <td width="8%" align="center"> <input name="SearchButton" type="submit" class="button01-out" value="确  定">
      </td>
    </tr>
  </form>
</table>
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
          <td>·名词说明：当前（搜索）</td>
        </tr>
        <tr>
          <td>　　只搜索当前所在资源类别下的所有资源（含子类的所有资源）</td>
        </tr>
      </table>
      <a name="Help"></a></td>
  </tr>
</table>
</body>
</html>
<%
Function ExeSql(user)

    Dim tSql,Strsql
    
    Select Case Work
        Case "Dustbin"        '回收站列表
           'tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=1 Order By Id DESC"
			'为了控制每个用户只能看到自己有权看的信息，作如下操作
			
			IF user="管理员" then
				tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=1 Order By Id DESC"
			Else
				Strsql="select ID from classlist where title='"&QXMC&"'"
				set Rs3=server.CreateObject ("Adodb.Recordset")
				Rs3.Open Strsql,conn,1,3		
				TypeID=Rs3("ID")		'查找第一级类别ID
				Rs3.Close
				
				IF parent="0" then			
					Parent=TypeID
				End IF
				If Column<>"" then
					Parent=GetClassID(QXMC)
				End if
				
				tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=1 Order By Id DESC"				
			End IF
        Case "UnChecked"    '未审核列表
            tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And IsChecked=0 And Del=0 Order By Id DESC"
        Case Else            '资源列表           
			'为了控制每个用户只能看到自己有权看的信息，作如下操作
			
			If user="管理员" then  '若是系统管理员
				'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And Content is Not null Order By Id DESC"
				tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And (Filename is null or Filename='') Order By Id DESC"
			Else '不是系统管理员
				Strsql="select ID from classlist where title='"&QXMC&"'"
				set Rs2=server.CreateObject ("Adodb.Recordset")
				Rs2.Open Strsql,conn,1,3		
					TypeID=Rs2("ID")		'查找第一级类别ID
				Rs2.Close
				Strsql2="select ID from classlist where Parent="&TypeID		'查找第二级类别ID
				Rs2.Open Strsql2,conn,1,3
				IF Rs2.EOF then			'若没有下一级，如下
					if purview ="管理者" then
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And Content is Not null And ClassTitle='"&QXMC&"' Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And (Filename is null or Filename='') And ClassTitle='"&QXMC&"' Order By Id DESC"
					else
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And Editor='"&UserName&"' And Content is Not null And ClassTitle='"&QXMC&"' Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And Editor='"&UserName&"' And (Filename is null or Filename='') And ClassTitle='"&QXMC&"' Order By Id DESC"
					end if
				Else					'若有下一级，如下
					IF parent="0" then			
						Parent=TypeID						
					End IF					
		 
					'这里又分为可以由分局来管的栏目和分局无权管的栏目
					If (purview = "管理者" and column="") then						
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Content is Not null Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And (Filename is null or Filename='') Order By Id DESC"																										
					Elseif(purview="管理者" and column<>"") then												
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Content is Not null Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And (Filename is null or Filename='') Order By Id DESC"
					Elseif (purview="编辑者" and column="") then						
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Editor='"&UserName&"' And Content is Not null Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Editor='"&UserName&"' And (Filename is null or Filename='') Order By Id DESC"										
					Elseif(purview="编辑者" and column<>"") then						
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Editor='"&UserName&"' And Content is Not null Order By Id DESC"						
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Editor='"&UserName&"' And (Filename is null or Filename='') Order By Id DESC"
					End If
				End IF
            End If
       
    End Select
 
    ExeSql=tSql
End Function

function GetClassID(QXMC)
	sql="select top 1 * from classlist where title='"&QXMC&"'"
	set cRs1=server.CreateObject ("Adodb.Recordset")	'查找第一级类别的ID号
	cRs1.Open sql,conn,1,3
	ClassID=cRs1("ID")	'第一级类别ID号
	cRs1.Close 
	sql="select top 1 * from classlist where title='"&column&"' and parent="&ClassID	'查找第二级类别号
	cRs1.Open sql,conn,1,3
	GetClassID=cRs1("ID")	
	cRs1.Close 
	set cRs1=nothing
End Function
%>