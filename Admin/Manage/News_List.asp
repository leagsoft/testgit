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
QXMC=Session("QXMC")	'�������
YHZL=Session("YHZL")	'�û����
UserName=Session("YHDL")	'�û�����

Purview = Session("Purview")	'�û�����Ȩ�޽�ɫ

Column = Session("Column")		'�ڶ����������
if Column="ʡ��" then Column="�㶫��ܾ�" end if
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

//������ɾ����Դ
//������������,�Ƿ���ʵɾ��
function DeleteReco(obj,eventObj,mRealDel)
{
	var selNum
		selNum=chkCheckBox(obj.Id)
	if(selNum==0)
	{
		alert("��ѡ����Ҫ[ɾ��/����ɾ��/�Ȼ�]����Դ")
		return false
	}
	if(confirm("��ȷ��Ҫ[ɾ��/����ɾ��/�Ȼ�]ѡ�еģ�"+selNum+"������Դ��"))
	{
		obj.Work.value="DelReco"
		obj.RealDel.value=mRealDel
		eventObj.disabled=true
		obj.submit()
	}
}
//�����������Դ����ʱ��ʹ�ô˹��ܣ�
//������������
function CheckReco(obj,eventObj)
{
    var selNum
        selNum=chkCheckBox(obj.Id)
    if(selNum==0)
    {
        alert("��ѡ����Ҫ[���]����Դ")
        return false
    }
    obj.Work.value="CheckReco"
    eventObj.disabled=true
    obj.submit()
}

//������������Դ����ʱ��ʹ�ô˹��ܣ�
//������������
function CreateFile(obj,eventObj)
{
    var selNum
        selNum=chkCheckBox(obj.Id)
    if(selNum==0)
    {
        alert("��ѡ����Ҫ[����]����Դ")
        return false
    }
    if(obj.Id.length)
    {
        for(var i=0;i<obj.Id.length;i++)
        {
            if(obj.Id[i].checked && obj.Id[i].HaveChecked!="True")
            {
                alert("��ǰ��ѡ����δ��ȫ[���]ͨ��")
                obj.Id[i].focus()
                return false
            }
        }
    }else{
        if(obj.Id.checked && obj.Id.HaveChecked!="True")
        {
            alert("��ǰ��ѡ����δ��ȫ[���]ͨ��")
            obj.Id.focus()
            return false
        }
    }
    obj.Work.value="CreateSelectedFile"
    eventObj.disabled=true
    obj.submit()
}

//�������༭��Դ
//������Button,��¼Id
function MdyReco(Id)
{
    event.srcElement.disabled=true
    document.body.innerHTML="<div align='center'>���Ե�,ϵͳ���ڴ����ݿ��ж�ȡ����Ҫ�༭����Դ��Ϣ...</div>"
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
          <td width="81%" bgcolor="#f6f6f6">��ǰλ�ã�<%=GetClassPath2(SysAdmin.AdminTopClassId,Parent,"")%></td>
          <td width="19%" align="right" bgcolor="#f6f6f6"><a href="#AdvanceSh">[����] 
            &nbsp; </a> </td>
        </tr>
        <tr> 
          <td colspan="2"> 
<%
Dim Sql,Rs
IF YHZL="����Ա" Then
            Sql="Select * From ClassList Where Parent="&Parent&" Order By UpTime"
Else		
'****ֻ��һ�����*******
		If Column="" then
			set Rs1=server.CreateObject ("Adodb.Recordset")
			Sql2="select * from ClassList where Title='"&QXMC&"'"
			
		
		
			Rs1.Open Sql2,conn,1,3
			if not rs1.eof then
			Parent=Rs1("ID")		
			Rs1.Close 
			set Rs1=nothing		
				end if
'****�м������*********
		Else
			if Parent<cClassID then
				Parent=cClassID
			end if
		End If
'****�����ߡ��༭�ߵĲ�ѯ���******		
			Sql="select * from classlist where Parent="&Parent
End IF            
            Set Rs=Conn.ExeCute(Sql)
            If Rs.Eof And Rs.Bof Then
                Response.Write("<font color='#666666'>�������</font>")
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
                'Response.Write "<font color=""#FF0000"">��Դ����վ�б�</font>&nbsp;&nbsp;[<A HREF=""#"" onclick=""if(confirm('�Ƿ�ȷ��Ҫ���[����վ]')){window.location='News_Mdy.asp?Work=ClearDustbin'}else{return false}""><font color=""#33CC00"">��ջ���վ</font></A>]"
                Response.Write "<font color=""#FF0000"">��Դ����վ�б�</font>"
            Case "UnChecked"
                Response.Write "<font color=""#FF0000"">δ�����Դ�б�</font>"      
            Case Else
                Response.Write "<font color=""#FF0000"">������Դ�б�</font>"
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
      <td width="50%">��Դ����</td>
    <td width="24%" align="center">���������</td>
    <td width="11%" align="center">���ʱ��</td>
    <td width="10%" align="center">�༭</td>
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
        <input type="checkbox" name="Id" value="<%=Rs("Id")%>" title="��¼��ţ�<%=Rs("Id")%>" HaveChecked="<%=CBool(Rs("IsChecked"))%>">
      </td>
      <td width="50%" bgcolor="#FFFFFF" title="չ���������" onclick="ShowControlPane(window.trNews_<%=Rs("Id")%>)"><label title="���α༭:<%=Rs("EditorTitle")%>"><%=Rs("title")%></label></td>
    <td width="24%" align="center" bgcolor="#FFFFFF"><%=Rs("ClassTitle")%></td>
    <td width="11%"  align="center" bgcolor="#FFFFFF" title="������:<%=FormatDateTime(Rs("UpTime"),1)%>"><%=FormatDateTime(Rs("AddTime"),2)%></td>
    <td width="10%" align="center" bgcolor="#FFFFFF">
        <input name="Submit2" type="button" class="button01-out" value="�� ��" onClick="MdyReco('<%=Rs("Id")%>')">
      </td>
  </tr>
  <tr Id="trNews_<%=Rs("Id")%>" <%If Not Def_ShowNewsContorlPlane Then Response.Write "style=""display:none""" End If%>>
  <td colspan="5" bgcolor="#FFFFFF">
<%
  If CBool(Rs("Del")) Then
    Response.Write "&nbsp;<a href=""News_Mdy.asp?Work=DelReco&Id="&Rs("Id")&""" title=""�ָ���Դ"">�ָ���Դ</a>"
    Response.Write "&nbsp;<a href=""News_Mdy.asp?Work=DelReco&Id="&Rs("Id")&"&RealDel=1"" title=""���ɻָ�"">����ɾ��</a>"
  Else
    Response.Write "&nbsp;<a href=""News_Mdy.asp?Work=DelReco&Id="&Rs("Id")&""" title=""�������վ"">ɾ����Դ</a>"
  End If
   'Response.Write "&nbsp;<a href=""Comment_List.asp?Work=ByNews&sType=ResId&sKey="&Rs("Id")&""">�鿴����</a>"
%>
  </td>
  </tr>
<%
      Rs.MoveNext
  Next
%>
<%If Rs.Eof And Rs.Bof Then%>
  <tr>
  <td align="center" colspan="7" bgcolor="#f6f6f6">������ؼ�¼</td>
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
        ��ѡ</label> 
        <%If Work="Dustbin" Then%>
        <input name="Submit3223" type="button" class="button01-out" value="��  ��" onClick="DeleteReco(form2,event.srcElement,0)" title="[�Ȼ�]��ǰ�Ѿ���[����ɾ��]�ļ�¼">
        <input name="Submit3222" type="button" class="button02-out" value="����ɾ��" onClick="DeleteReco(form2,event.srcElement,1)" title="[����ɾ��]����Դ���޷���ԭ">
        <%Else%>
        <input name="Submit222" type="button" class="button01-out" value="ɾ ��" title="����ɾ��������վ" onClick="DeleteReco(form2,event.srcElement,0)">
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
        ����: 
        <select name="sType" class="Input">
          <option value="Title" <%If sType="Title" Then Response.Write("selected") End If%>>��Դ����</option>
          <option value="ClassTitle" <%If sType="ClassTitle" Then Response.Write("selected") End If%>>��Դ����</option>
          <option value="Content" <%If sType="Content" Then Response.Write("selected") End If%>>��Դ����</option>
          <option value="AuthorTitle" <%If sType="AuthorTitle" Then Response.Write("selected") End If%>>��������</option>
          <option value="FromTitle" <%If sType="FromTitle" Then Response.Write("selected") End If%>>��Դ��Դ</option>          
        </select> <input name="Parent" type="radio" class="Input" value="0" checked>
        ���� 
        <input name="Parent" type="radio" class="Input" value="<%=Parent%>">
        ��ǰ </td>
      <td width="25%" align="right"> <input name="sKey" type="text" class="Input" id="sKey" style="width:100%" value="<%=Trim(Request("sKey"))%>"></td>
      <td width="8%" align="center"> <input name="SearchButton" type="submit" class="button01-out" value="ȷ  ��">
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
          <td>������˵������ǰ��������</td>
        </tr>
        <tr>
          <td>����ֻ������ǰ������Դ����µ�������Դ���������������Դ��</td>
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
        Case "Dustbin"        '����վ�б�
           'tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=1 Order By Id DESC"
			'Ϊ�˿���ÿ���û�ֻ�ܿ����Լ���Ȩ������Ϣ�������²���
			
			IF user="����Ա" then
				tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=1 Order By Id DESC"
			Else
				Strsql="select ID from classlist where title='"&QXMC&"'"
				set Rs3=server.CreateObject ("Adodb.Recordset")
				Rs3.Open Strsql,conn,1,3		
				TypeID=Rs3("ID")		'���ҵ�һ�����ID
				Rs3.Close
				
				IF parent="0" then			
					Parent=TypeID
				End IF
				If Column<>"" then
					Parent=GetClassID(QXMC)
				End if
				
				tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=1 Order By Id DESC"				
			End IF
        Case "UnChecked"    'δ����б�
            tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And IsChecked=0 And Del=0 Order By Id DESC"
        Case Else            '��Դ�б�           
			'Ϊ�˿���ÿ���û�ֻ�ܿ����Լ���Ȩ������Ϣ�������²���
			
			If user="����Ա" then  '����ϵͳ����Ա
				'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And Content is Not null Order By Id DESC"
				tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And (Filename is null or Filename='') Order By Id DESC"
			Else '����ϵͳ����Ա
				Strsql="select ID from classlist where title='"&QXMC&"'"
				set Rs2=server.CreateObject ("Adodb.Recordset")
				Rs2.Open Strsql,conn,1,3		
					TypeID=Rs2("ID")		'���ҵ�һ�����ID
				Rs2.Close
				Strsql2="select ID from classlist where Parent="&TypeID		'���ҵڶ������ID
				Rs2.Open Strsql2,conn,1,3
				IF Rs2.EOF then			'��û����һ��������
					if purview ="������" then
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And Content is Not null And ClassTitle='"&QXMC&"' Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And (Filename is null or Filename='') And ClassTitle='"&QXMC&"' Order By Id DESC"
					else
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And Editor='"&UserName&"' And Content is Not null And ClassTitle='"&QXMC&"' Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent) &") And Del=0 And Editor='"&UserName&"' And (Filename is null or Filename='') And ClassTitle='"&QXMC&"' Order By Id DESC"
					end if
				Else					'������һ��������
					IF parent="0" then			
						Parent=TypeID						
					End IF					
		 
					'�����ַ�Ϊ�����ɷ־����ܵ���Ŀ�ͷ־���Ȩ�ܵ���Ŀ
					If (purview = "������" and column="") then						
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Content is Not null Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And (Filename is null or Filename='') Order By Id DESC"																										
					Elseif(purview="������" and column<>"") then												
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Content is Not null Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And (Filename is null or Filename='') Order By Id DESC"
					Elseif (purview="�༭��" and column="") then						
						'20050630 tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Editor='"&UserName&"' And Content is Not null Order By Id DESC"
						tSql="SELECT Id,Class,Title,EditorTitle,AddTime,UpTime,ClassTitle,Del,Created,IsChecked,FilePath From view_AllNewsInfo Where "&sType&" Like '%"&sKey&"%' And Class In ("& Parent&AllChildClass(Parent)&") And Del=0 And Editor='"&UserName&"' And (Filename is null or Filename='') Order By Id DESC"										
					Elseif(purview="�༭��" and column<>"") then						
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
	set cRs1=server.CreateObject ("Adodb.Recordset")	'���ҵ�һ������ID��
	cRs1.Open sql,conn,1,3
	ClassID=cRs1("ID")	'��һ�����ID��
	cRs1.Close 
	sql="select top 1 * from classlist where title='"&column&"' and parent="&ClassID	'���ҵڶ�������
	cRs1.Open sql,conn,1,3
	GetClassID=cRs1("ID")	
	cRs1.Close 
	set cRs1=nothing
End Function
%>