<%
'����˵�
cNav = "Prdct"
%>
<!--#include file="Head.asp"-->
<!--#include File="Include/Conn.asp"-->
<!--#include File="Include/Function.asp"-->
<%
'���尴ť
cAdd    = "���Ӽ�¼"
cSave   = "�����¼"
cDel    = "ɾ����¼"
cList   = "�б���ʾ"

'���հ���
cAction = Trim(Request("cAction"))
cMode   = Trim(Request("cMode"))

'���ձ���
nProId      = Request("nProId")
nType1      = Replace(Trim(Request("nType1")),"'","��")
nProName    = Replace(Trim(Request("nProName")),"'","��")
nIntroduct  = Replace(Trim(Request("nIntroduct")),"'","��")
nPic        = Replace(Trim(Request("nPic")),"'","��")

If nProId = "" or cAction=cAdd Then
   cMode = "New"
   cMsg  = "¼���µĲ�Ʒ���ݣ�"
Else 
   cMode = "Edit"
End If

'�жϰ���
If cAction = cSave Then
       '�����ж�
       If nType1 = "" Then
          cMsg  = "��ѡ���Ʒ���࣡"
       ElseIf nProName = "" Then
	  cMsg  = "��Ʒ���Ʋ���Ϊ�գ�"
       ElseIf cMode = "New" Then
          'д�����ݿ�
           cSql="insert into Products(Type1,Proname,Content,Pic,CreDate) values ('"&nType1&"','"&nProName&"','"&nIntroduct&"','"&nPic&"',getdate())"
	  On Error Resume Next
	  Conn.Execute(cSql)
	  If Err.Number = 0 Then
		 cMsg="��¼���ӳɹ���"
		 Session("cApply")=""
		 cMode = "Edit"
	  Else
		 cMsg = "���ݴ���ʧ�ܣ�"
		 cMode = "New"
	  End If
	  Err.Clear
       ElseIf cMode = "Edit" Then
          'д�����ݿ�
		cSql="Update Products set Type1='"&nType1&"',Proname='"&nProname&"',Content='"&nIntroduct&"',Pic='"&nPic&"'"&" where PROID="&nProid
		'Response.Write csql
		'Response.End 
	  On Error Resume Next	 
	  Conn.Execute(cSql)
	  If Err.Number = 0 Then
		 cMsg = "��¼���ĳɹ���"
		 Session("cApply")=""
	  Else
		 cMsg = "���ݴ���ʧ�ܣ�"
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
          cMsg = "���ݴ���ʧ�ܣ�"
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
   '��ѯ���ݿ�
   cSql = "select * from Products where PROID = "&nProId
   Set Rs = Conn.Execute(cSql)
   If Rs.Eof Then
      Response.Redirect "Index.asp"
   Else
      '�������
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
    alert("��ѡ��ͼƬ���");
  theForm.nType1.focus();
  return(false);
  }
  if (theForm.nProName.value==null)
  {
    alert("������ͼƬ���⣡");
	theForm.nProName.focus();
	return(false);
  }
  if (theForm.nPic.value==null)
  {
    alert("���ϴ�ͼƬ��");
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
    <td colspan="2" bgcolor="#FFDA99" valign="middle" height="30">&nbsp;<a href="index.asp">ͼƬ����</a> -> <font color=red>ͼƬ��Ϣά��</font></td>
  </tr>
  <tr>
    <td colspan="2" height="22" bgcolor="#00DDDD" valign="middle" nowrap>&nbsp;��ʾ��Ϣ��<%= cMsg %></td>
  </tr>
  <tr>
    <td height="221" align="center"> 
<table width="95%" border=0 cellpadding=1 cellspacing=0 align="center">
    <tr>
      <td width="20%" align="right"><font color="brown">ͼƬ���ࣺ</td>
      <td width="50%">
<!--<select name="nType1" onchange="javascript:document.PrdctForm.submit()">-->
<select name="nType1">
<option value="">��ѡ������</option>
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
      <td width="20%" align="right"><font color="brown">ͼƬ���⣺</td>
      <td width="80%"><input type="text" name="nProName" size="50" maxlength="100" value="<%=nProName%>"></td>
    </tr>
    <tr>
      <td width="20%" align="right"><font color="brown">ͼƬ���ܣ�</td>
      <td width="80%"><textarea rows="10" name="nIntroduct" cols="50" wrap><%=nIntroduct%></textarea></td>
    </tr>
    <tr>
      <td width="20%" align="right"><font color="brown">ͼƬ��</td>
      <td width="80%"><INPUT TYPE="Text" Name="nPic" Size="50" value="<%=nPic%>"> <A onClick="window.open('FileSystem/View.asp','ResWin','resizable,scrollbars,width=600,height=500')" style="cursor:hand" title="��������ͼƬ"> <font color=red><u>����</u></font> </A></td>
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


