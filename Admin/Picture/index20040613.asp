<%
'����˵�
cNav = "Prdct"
%>
<!--#include file="Head.asp"-->
<!--#include File="Include/Conn.asp"-->
<!--#include File="Include/Function.asp"-->
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

'���尴ť
cFirst = "��ǰһҳ"
cPrev = "��һҳ"
cNext = "��һҳ"
cLast = "���һҳ"
cSearch = "��ѯ"
cAdd    = "���Ӽ�¼"
cLook = "�鿴ѯ�۵�"

'���巭ҳ����
pnRecPerPage = 10
pnRecordCount = 0
pnRecordRest = 0
pnPageCount = 0
pnCounter = 0
pnCurrentPage = 0

'���ղ�ѯ����
cType       = Trim(Request("cType"))
precType    = Trim(request("precType"))
cParentType = Trim(Request("cParentType"))
precParentType=Trim(Request("precParentType"))
cPro        = Trim(Request("cPro"))
cKeyword    = Trim(Replace(Request("cKeyword"),"'",""))
cPriceFrom  = Trim(Replace(Request("cPriceFrom"),"'",""))
cPriceTo    = Trim(Replace(Request("cPriceTo"),"'",""))
cSort       = Trim(Replace(Request("cSort"),"'",""))

'�����������
If cType = "" Then
   cSql3     = ""
   cTyepSql1 = ""
   cTitle    = "ͼƬ����"
Else
   cSql3     = "where TYPE1="&cType
   cTyepSql1 = " and (PARENTID="&cType&")"
   'cTitle    = "<a href=PrdctList.asp><font color=black><u>"&cParentType&"</u></font></a>"
   'cTitle    = "<a href=PrdctList.asp?cType="&precType&"&cParentType="&precParentType&"><font color=black><u>"&cParentType&"</u></font></a>"
End If
If cPro = "" Then
   cSql1 = ""
ElseIf cPro = "1" Then
   cSql1 = " and (ALLOWSYS='')"
ElseIf cPro = "2" Then
   cSql1 = " and (ALLOWSYS<>'')"   
End If
If cKeyword = "" Then
   cSql2 = ""
Else
   cSql2 = " and (PRONAME like '%"+ cKeyword +"%' or PROCODE like '%"+ cKeyword +"%')"
End If
If cPriceFrom = "" and cPriceTo = "" Then
   cSql4 = ""
ElseIf cPriceFrom <> "" and cPriceTo = "" Then
   cSql4 = " and (convert(money,PRICE)>"&cPriceFrom&")"
ElseIf cPriceFrom = "" and cPriceTo <> "" Then
   cSql4 = " and (convert(money,PRICE)<"&cPriceTo&")"
ElseIf cPriceFrom <> "" and cPriceTo <> "" Then
   cSql4 = " and (convert(money,PRICE) between "&cPriceFrom&" and "&cPriceTo&")"
End If
If cSort = "" Then
   cSql5 = " order by CreDATE desc"
ElseIf cSort = "1" Then
   cSql5 = " order by CreDATE asc"   
ElseIf cSort = "2" Then
   cSql5 = " order by convert(money,PRICE) desc"
ElseIf cSort = "3" Then
   cSql5 = " order by convert(money,PRICE) asc"
End If  

'Execute a query
cSql = "select * from Products "&cSql1&cSql4&cSql2&cSql3&cSql5
'Response.Write cSql

'Response.Write csql
'Response.End 
'Response.Write cSql
'Set Rs = mydb.Execute(cSql)
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
<td width="50%" align="left" class="BigFont"><img src="Images/item.gif" border="0" width="18" height="18" align="absmiddle">&nbsp;</td>
<td width="50%" align="right"><a href="TypeAdmin.asp" onclick="return js_t(this.href);" title="����ͼƬ�ķ���"><span class="BigFont"><font color="red"><u>�������<u></font></span></a></td>
</tr>
<tr>
<td width="100%" colspan="2" bgcolor="#E1E1E1" height="1"></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
  <table border="0" cellspacing="1" width="100%">
    <tr>
<%
'��ѯ��Դ����
TypeSql = "select DICID,Type from SYSDIC where DELETED=0 order by DICID asc"
'Response.Write Typesql
'response.write TypeSql
'Set RsSub = mydb.Execute(TypeSql)
Set RsSub = Conn.Execute(TypeSql)

If Not RsSub.Eof Then
i = 1
Do 
  If RsSub.Eof Then Exit Do
  j = i Mod 4
  'iCount=GetCount("select PROID from PRODUCTS where TYPE1="&RsSub("DICID")&" or TYPE2="&RsSub("DICID")&" or TYPE3="&RsSub("DICID"))
%>
<td width="25%" valign="top" height="20"><a href="index.asp?cType=<%= RsSub("DICID")%>&cParentType=<%=RsSub("Type")%>&precType=<%=cType%>&precParentType=<%=cParentType%>"><u><%If Trim(RsSub("DICID"))=cType Then Response.Write "<font color=black>" End If%><%= Trim(RsSub("Type"))%></u></a></td>
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
<form name="PrdctList" method="post" action="index.asp?cType=<%=cType%>&cParentType=<%=cParentType%>">
<table width="80%" border="1" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFDFDF" bordercolorlight="#333333" bordercolordark="#FFFFFF">
  <tr>
    <td colspan="4" bgcolor="#FFDA99" valign="middle" height="30">&nbsp;<a href="PrdctList.asp">ͼƬ����</a> -> <font color=red>ͼƬ�б�</font>&nbsp;&nbsp;<font color=red><%=cMsg%></font></td>
  </tr>
  <tr>
      <td bgcolor="#00DDDD" height="22" align="center">��Ʒ���</td>
      <td bgcolor="#00DDDD" height="22" align="center">��Ʒ����</td>
      <!--<td bgcolor="#00DDDD" height="22" align="center">�۸�</td>-->
      <td bgcolor="#00DDDD" nowrap height="22" align="center" colspan="2">��������</td>
      <!--<td bgcolor="#00DDDD" nowrap height="22" align="center">�����</td>-->
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
      <td height="22" align="center" colspan="2"><%= Rs("CREDATE") %></td>  
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
      ��<font color="BROWN"><%= pnRecordCount %></font>����¼ ��<font COLOR="BROWN"><%= fnStartRecord %></font>-<font COLOR="BROWN"><%= fnEndRecord %></font>�� ��<font COLOR="BROWN"><%= pnCurrentPage %></font>/<font COLOR="BROWN"><%= pnPageCount %></font>ҳ
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
'mydb.Close
'Set mydb = Nothing
Conn.Close
Set Conn = Nothing
%>
<input TYPE="button" NAME="cAction" VALUE="<%= cAdd %>" onclick="javascript:location.href='PrdctForm.asp'">
<!--<input TYPE="button" NAME="cAction" VALUE="<%= cLook %>" onclick="javascript:location.href='ENQUIRYList.asp'">-->
    </td>
    </form>
  </tr>
</table>
<br>
<br>
<!--#include file="Foot.asp"-->

