<%
'定义菜单
cNav = "Prdct"
%>
<!--#include file="Include/conn.asp"-->
<!--#include file="Head.asp"-->
<%
cAction = Trim(Request("cAction"))
cMsg    = Trim(Request("cMsg"))
If IsNull(Trim(Request("fnStartRecord"))) or Trim(Request("fnStartRecord")) = "" Then
	fnStartRecord = 1
Else
	fnStartRecord = CInt(Trim(Request("fnStartRecord")))
End If

If IsNull(Trim(Request("fnEndRecord"))) or Trim(Request("fnEndRecord")) = "" Then
	fnEndRecord = 1
Else
	fnEndRecord = CInt(Trim(Request("fnEndRecord")))
End If

'定义按钮
cFirst = "最前一页"
cPrev = "上一页"
cNext = "下一页"
cLast = "最后一页"
cSearch = "查询"
cAdd    = "增加记录"
cLook = "查看询价单"

'定义翻页变量
pnRecPerPage = 10
pnRecordCount = 0
pnRecordRest = 0
pnPageCount = 0
pnCounter = 0
pnCurrentPage = 0

'接收查询条件
cType       = Trim(Request("cType"))

'生成查询条件
If cType = "" Then
	Csql="Select * from Products"
Else
	Csql="Select * from Products where Type1="&cType
End If

Set Rs = Conn.Execute(cSql)

Do
	If Rs.Eof Then Exit Do
	pnRecordCount = pnRecordCount + 1
	Rs.MoveNext
	Loop
pnPageCount = Int(pnRecordCount / pnRecPerPage)
pnRecordRest = pnRecordCount - pnPageCount * pnRecPerPage
If pnRecordRest <> 0 Then
	pnPageCount = pnPageCount + 1
End If

If cAction = cFirst Then
	fnStartRecord = 1
ElseIf cAction = cPrev Then
	fnStartRecord = fnStartRecord - pnRecPerPage
	If fnStartRecord <= 0 Then
		fnStartRecord = 1
	End If
ElseIf cAction = cNext Then
	fnStartRecord = fnEndRecord + 1
	If fnStartRecord > pnRecordCount Then
		fnStartRecord = pnRecordCount
	End If
ElseIf cAction = cLast Then
	If pnRecordRest > 0 Then
		fnStartRecord = pnRecordCount - pnRecordRest + 1
	Else
		fnStartRecord = pnRecordCount - pnRecPerPage + 1
	End If
	If fnStartRecord <= 0 Then
		fnStartRecord = 1
	End If
End If
pnCurrentPage = Int(fnStartRecord / pnRecPerPage)
If pnCurrentPage <> fnStartRecord / pnRecPerPage Then
	pnCurrentPage = pnCurrentPage + 1
End If
%>
<table border="0" cellspacing="1" width="80%" align="center">
<tr>
<td width="50%" align="left" class="BigFont"><img src="/Images/item.gif" border="0" width="18" height="18" align="absmiddle">&nbsp;图片分类</td>
<td width="50%" align="right"><a href="" onclick="return js_t(this.href);" title="按此产品的分类"><span class="BigFont"><font color="red"><u>分类管理<u></font></span></a></td>
</tr>
<tr>
<td width="100%" colspan="2" bgcolor="#E1E1E1" height="1"></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
  <table border="0" cellspacing="1" width="100%">
    <tr>
<%
'查询资源分类
TypeSql = "select DICID from SYSDIC where DELETED=0 order by DICID asc"
'response.write TypeSql
Set RsSub = Conn.Execute(TypeSql)

If Not RsSub.Eof Then
i = 1
Do 
  If RsSub.Eof Then Exit Do
  j = i Mod 4
%>
<td width="25%" valign="top" height="20"><a href="PrdctList.asp?cType=<%= RsSub("DICID")%>"><u><%If Trim(RsSub("DICID"))=cType Then Response.Write "<font color=black>" End If%><%= Trim(RsSub("Type"))%></u></a></td>
<%
  RsSub.MoveNext
  If (j=0) and (Not RsSub.Eof) Then
%>
</tr>
<tr>
<%
  Else
      If (RsSub.Eof) and (j>0) then
         For k=j+1 to 3
%>
<td></td>
<%
         Next
%>
</tr>
<%
      End If
  End if
i = i + 1
Loop
End If
RsSub.Close
Set RsSub = Nothing
%>
<%
  RsSub.MoveNext
  If (j=0) and (Not RsSub.Eof) Then
%>
</tr>
<tr>
<%
  Else
      If (RsSub.Eof) and (j>0) then
         For k=j+1 to 3
%>
<td></td>
<%
         Next
%>
</tr>
<%
      End If
  End if
i = i + 1
Loop
End If
RsSub.Close
Set RsSub = Nothing
%>
 </tr>
</table>
</td>
</tr>
</table>
<form name="PrdctList" method="post" action="PrdctList.asp?cType=<%=cType%>">
<table width="80%" border="1" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFDFDF" bordercolorlight="#333333" bordercolordark="#FFFFFF">
  <tr>
    <td colspan="4" bgcolor="#FFDA99" valign="middle" height="30">&nbsp;<a href="PrdctList.asp">产品管理</a> -> <font color=red>产品列表</font>&nbsp;&nbsp;<font color=red><%=cMsg%></font></td>
  </tr>
  <tr>
    <td colspan="4" BGCOLOR="teal" valign="middle" height="30" align="center">
<font color=white>关键字：<input type="text" name="cKeyword" size="20" style="background-color:#FFB015;" value="<%= cKeyword%>">
<select name="cPro" style="background-color:#FFB015;">
<option value="">所有产品</option>
<option value="1" <% If cPro = "1" Then Response.Write "selected" End If%>>公开产品</option>
<option value="2" <% If cPro = "2" Then Response.Write "selected" End If%>>非公开产品</option>
</select>
<input type="submit" NAME="cAction" value="<%= cSearch%>">
    </td>
  </tr>
  <tr>
      <td bgcolor="#00DDDD" height="22" align="center">产品编号</td>
      <td bgcolor="#00DDDD" height="22" align="center">产品名称</td>
      <!--<td bgcolor="#00DDDD" height="22" align="center">价格</td>-->
      <td bgcolor="#00DDDD" nowrap height="22" align="center" colspan="2">发布日期</td>
      <!--<td bgcolor="#00DDDD" nowrap height="22" align="center">浏览数</td>-->
  </tr>
<%          
If pnRecordCount > 0 Then
	Rs.MoveFirst
	Rs.Move fnStartRecord - 1
	Do
		If Rs.Eof Then Exit Do
		
%>		
  <tr>
      <td height="22"><%= Replace(Trim(Rs("PROID")),cKeyword,"<font color=red>"&cKeyword&"</font>")%></td>
      <td height="22"><a href="PrdctForm.asp?nProId=<%=Rs("PROID")%>"><%= Replace(Trim(Rs("PRONAME")),cKeyword,"<font color=red>"&cKeyword&"</font>")%></a></td>
      <!--<td height="22" align="right"><%If Rs("Price")="" Then Response.Write "面议" Else Response.Write Trim(Rs("CUnit"))&" "&Rs("Price") End If%></td>--> 
      <td height="22" align="center" colspan="2"><%= Rs("CreDATE") %></td>  
      <!--<td align="center" height="22"><%If Rs("HITCOUNT")>100 Then Response.Write "<font color=red>"%><%= Rs("HITCOUNT") %></td>-->
  </tr>
	<%
		pnCounter = pnCounter + 1
		If pnCounter >= pnRecPerPage Then Exit Do
		Rs.MoveNext
		Loop
	fnEndRecord = fnStartRecord + pnCounter - 1
End If
%>
  <tr>
      <td colspan="4" align="center" bgcolor="#00DDDD" height="22">
      共<font color="BROWN"><%= pnRecordCount %></font>条记录 第<font COLOR="BROWN"><%= fnStartRecord %></font>-<font COLOR="BROWN"><%= fnEndRecord %></font>条 第<font COLOR="BROWN"><%= pnCurrentPage %></font>/<font COLOR="BROWN"><%= pnPageCount %></font>页
      </td>
  </tr>
  <tr>
    <td colspan="4" bgcolor="#FFDA99" height="30" align="center" valign="middle">
	<input TYPE="Hidden" NAME="fnStartRecord" VALUE="<%= fnStartRecord %>">
	<input TYPE="Hidden" NAME="fnEndRecord" VALUE="<%= fnEndRecord %>">
	<input TYPE="Hidden" NAME="pnRecordCount" VALUE="<%= pnRecordCount %>">
<%
If fnStartRecord > 1 Then
%>
	<input TYPE="Submit" NAME="cAction" VALUE="<%= cFirst %>">
	<input TYPE="Submit" NAME="cAction" VALUE="<%= cPrev %>">
<%
End If
If fnEndRecord < pnRecordCount Then
%>
	<input TYPE="Submit" NAME="cAction" VALUE="<%= cNext %>">
	<input TYPE="Submit" NAME="cAction" VALUE="<%= cLast %>">
<%
End If
Rs.Close
Set Rs = Nothing
mydb.Close
Set mydb = Nothing
%>
<input TYPE="button" NAME="cAction" VALUE="<%= cAdd %>" onclick="javascript:location.href='PrdctForm.asp'">
<input TYPE="button" NAME="cAction" VALUE="<%= cLook %>" onclick="javascript:location.href='ENQUIRYList.asp'">
    </td>
    </form>
  </tr>
</table>
<br>
<br>
<!--#include file="Foot.asp"-->


