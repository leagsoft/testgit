<!--#include file="../Include/Conn.asp" -->
<!-- #include file="../Include/ClassList_Fun.asp" -->
<!--#include file="../Include/Config.asp" -->
<!--#include file="../Include/Tkl_StringClass.asp" -->
<!--#include file="../Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="../Include/OnlineClass.asp" -->
<!--#Include File="../Include/UpdateAdminTime.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class

Call UpdateAdminTime()

Dim cFun
Set cFun=New Tkl_StringClass
QXMC=Session("QXMC")		'ȡ��Ȩ������		'Add benny
Column=Session("Column")	'�û����Ƶ���Ŀ
if column="ʡ��" then
	Column="�㶫��ܾ�"
end if
%>
<html>
<head>
<title>News_Add.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="../Include/ManageStyle.css" type="text/css">
<script language="JavaScript" src="../Include/Tkl_ClassTree.js" type="text/JavaScript"></script>
<script src="../Include/Tkl_Skin.js"></script>
</head>
<body bgcolor="#FFFFFF" leftmargin="5" topmargin="5">
<script src="../Include/Tkl_Tooltip.js"></script>
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
      <td colspan="2" class="BarTitleBg"> �����Դ </td>
    </tr>
    <tr> 
      <td width="16%" height="9" valign="top" class="BarTitle">��Դ����:</td>
      <td width="84%" height="9" bgcolor="#FFFFFF"> 
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="68%">
<%if Session("YHZL")<>"����Ա" then%>    
    <script language="javascript">
      var root1
      root1=CreateRoot("myTree1","����ѡ��[��Դ���]")
      <%
	  //Dim CurrentClassIdUsed
		  //CurrentClassIdUsed=Request.Cookies("ZGW_NewsSys")("CurrentClassIdUsed")
		  CurrentClassIdUsed=Request.Cookies("ZGW_NewsSys3")("CurrentClassIdUsed")
	  If Not IsNumeric(CurrentClassIdUsed) Then
		CurrentClassIdUsed=-1
	  End If	
	  //Call CreateClassTree1(SysAdmin.AdminTopClassId,CLng(CurrentClassIdUsed))
	  Call CreateClassTree1(QXMC,Column,CLng(CurrentClassIdUsed))
	  %>
      </script>
<%elseif Session("YHZL")="����Ա" then%>      
    <script language="javascript">
      var root1
      root1=CreateRoot("myTree1","����ѡ��[��Դ���]")
      <%
	  //Dim CurrentClassIdUsed
		  CurrentClassIdUsed=Request.Cookies("ZGW_NewsSys3")("CurrentClassIdUsed")
	  If Not IsNumeric(CurrentClassIdUsed) Then
		CurrentClassIdUsed=-1
	  End If	
	  Call CreateClassTree3(SysAdmin.AdminTopClassId,CLng(CurrentClassIdUsed))
	  %>
      </script>
<%end if%> 
	</td>
    <td width="32%" align="right" valign="top"><font color="red">*</font>
	<label for="CurrentClassIdUsed" title="���������Դʱ�Զ�ʹ�õ�ǰ���">
	<input type="checkbox" id="CurrentClassIdUsed" name="CurrentClassIdUsed" value="1" <%If CurrentClassIdUsed<>-1 Then Response.Write "checked" End If%>>�´�ʹ��</label>	</td>
  </tr>
</table></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">��Դ����:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="Title" type="text" class="Input" id="Title" size="60">&nbsp;&nbsp;<font color="red">*</font> 
      </td>
    </tr>
<%	
	If Session("QXMC")="�ֳ�ר��" then
'������оֳ�������
		Strsql="select YHDL,UserRankRight from YHXX where YHBM='���쵼' order by UserRankRight"
		set cRs=server.CreateObject ("Adodb.Recordset")
			cRs.Open strsql,connect,1,3
%>
    <tr> 
      <td width="16%" class="BarTitle">��Դ����:</td>
      <td width="84%" bgcolor="#FFFFFF">
		<select name="Author">
<%				
			while not cRs.EOF 
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
      &nbsp;&nbsp;<font color="red">*</font> 
      </td>
    </tr>
<%	End If%>
<!--Add By BennyLiu:20040712-->
<%If Session("QXMC")="�־ֶ�̬" then%>
    <tr> 
      <td width="16%" class="BarTitle">�����:</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../../SetPurview/info.asp?page=30&type=group&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../../SetPurview/info.asp?page=30&type=user&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../../SetPurview/infogroup.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../../SetPurview/infouser.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Session("cBrowser")%>">
      </td>
    </tr>  
<%End If%>    
<!--End Add-->
<%If Session("QXMC")="����ͳ����Ϣ" then%>
    <tr> 
      <td width="16%" class="BarTitle">�����:</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../../SetPurview/info.asp?page=30&type=group&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../../SetPurview/info.asp?page=30&type=user&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../../SetPurview/infogroup.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../../SetPurview/infouser.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Session("cBrowser")%>">
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">�ļ����ͣ�</td>
      <td bgcolor="#FFFFFF">
		<select name="DocumentType">
			<option value="0">����</option>
			<option value="1">�ĵ�</option>
		</select>
      </td>
    </tr>    
<%End If%>
    <tr> 
      <td width="16%" class="BarTitle">�ϴ��ĵ�:</td>
      <!--<td bgcolor="#FFFFFF"><input type="submit" name="SaveAddButton" id="SaveAddButton" value="..." ></td>-->
      <td bgcolor="#FFFFFF">
		<input name="Url" type="text" readonly class="Input" id="Url" size="60" value="">
		<input type="button" name="Upload" value="�ϴ��ĵ�" onClick="javascript:receive();">
	  </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="13"> 
        <script>
function receive()
{
	//alert("aaa");
	//window.open("FileSystem/UpFile.asp"); 
	var result=window.showModalDialog ("FileSystem/UpFile.asp","","dialogWidth:500px;dialogHeight:230px;center:yes;scroll:no;");
		if (result!=false && result!=undefined)
		window.form1.Url.value=result;
}        
        
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
        alert("��ѡ��[��Դ���]")
        return false
    }
    if(obj.Title.value==""){
        alert("������[��Դ����]");
        obj.Title.focus();
        return false;
    }
    if(obj.Url.value==""){
            alert("�ĵ�����Ϊ��");
            obj.Url.focus();
            return false;
        }
    obj.SaveAddButton.disabled=true
    return true;
}
</script> </td>
      <td>
        <input type="submit" name="SaveAddButton" id="SaveAddButton" class="button01-out" value="ȷ����">&nbsp;&nbsp;<input name="Submit3" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
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
%>
<form name="form2" method="post" action="News_Mdy.asp?Work=SaveMdy" onsubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr align="center"> 
      <td colspan="2" class="BarTitleBg"> �༭��Դ 
        <input name="Id" type="hidden" id="Id" value="<%=Rs("Id")%>"></td>
    </tr>
    <tr> 
      <td width="16%" height="9" valign="top" class="BarTitle">��Դ����:</td>
      <td width="84%" height="9" bgcolor="#FFFFFF"> 
<%if Session("YHZL")<>"����Ա" then%>      
        <script>
      var root2
      root2=CreateRoot("myTree2","����ѡ��[��Դ���]")
      <%
		//Call CreateClassTree2(SysAdmin.AdminTopClassId,Rs("Class"))
		Call CreateClassTree2(QXMC,Column,Rs("Class"))
		%>
      </script>
<%elseif Session("YHZL")="����Ա" then%>      
        <script>
      var root2
      root2=CreateRoot("myTree2","����ѡ��[��Դ���]")
      <%
		Call CreateClassTree4(SysAdmin.AdminTopClassId,Rs("Class"))
	  %>
      </script>
<%end if%>   
<font color="red">*</font> 
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">��Դ����:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="Title" type="text" class="Input" id="Title" size="60" value="<%=cFun.HTMLEncode2(Rs("Title"))%>">&nbsp;&nbsp;<font color="red">*</font>  
      </td>
    </tr>
<%	
	If Session("QXMC")="�ֳ�ר��" then
'������оֳ�������
		Strsql="select YHDL,UserRankRight from YHXX where YHBM='���쵼' order by UserRankRight"
		set cRs2=server.CreateObject ("Adodb.Recordset")
			cRs2.Open strsql,connect,1,3
%>
    <tr> 
      <td width="16%" class="BarTitle">��Դ����:</td>
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
		</select> 
      &nbsp;&nbsp;<font color="red">*</font>  
      </td>
    </tr>
<%	End If%>
<!--Add By BennyLiu:20040712-->
<%If Session("QXMC")="�־ֶ�̬" then%>
    <tr> 
      <td width="16%" class="BarTitle">�����:</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../../SetPurview/info.asp?page=30&type=group&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../../SetPurview/info.asp?page=30&type=user&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Rs("Browser")%>">
      </td>
    </tr>  
<%End If%>    
<!--End Add-->
<!--Add By BennyLiu:20040625-->
<%If Session("QXMC")="����ͳ����Ϣ" then%>
    <tr> 
      <td width="16%" class="BarTitle">�����:</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../../SetPurview/info.asp?page=30&type=group&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../../SetPurview/info.asp?page=30&type=user&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Rs("Browser")%>">
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">�ļ����ͣ�</td>
      <td bgcolor="#FFFFFF">
		<select name="DocumentType">
			<option value="0"<%if Rs("IsDocument")="0" then Response.Write " selected" end if%>>����</option>
			<option value="1"<%if Rs("IsDocument")="1" then Response.Write " selected" end if%>>�ĵ�</option>
		</select>
      </td>
    </tr>    
<%End If%>    
<!--End Add-->
    <tr> 
      <td class="BarTitle">�ϴ��ĵ�:</td>
      <!--<td bgcolor="#FFFFFF"><input name="SmallImg" type="text" class="Input" id="SmallImg" value="<%=cFun.HTMLEncode2(Rs("Filename"))%>" size="50">
        <input type="button" name="Button3" value="..." onClick="window.open('FileSystem2/View.asp','ResWin','resizable,scrollbars,width=600,height=500')"></td>-->
      <td bgcolor="#FFFFFF">
		<input name="Url" type="text" class="Input" id="Url" size="60" value="<%=cFun.HTMLEncode2(Rs("FileName"))%>">
		<input type="button" name="Upload" value="�ϴ��ĵ�" onClick="javascript:receive2();">   
	  </td>        
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="27"> 
        <script>
function receive2()
{
	var result=window.showModalDialog ("FileSystem/UpFile.asp","","dialogWidth:500px;dialogHeight:230px;center:yes;scroll:no;");
	if (result!=false && result!=undefined)
		window.form2.Url.value=result;
}        
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
        alert("��ѡ��[��Դ���]")
        return false
    }
    if(obj.Title.value==""){
        alert("������[��Դ����]");
        obj.Title.focus();
        return false;
    }
    if(obj.Url.value==""){
            alert("�ĵ�����Ϊ�գ�");
            obj.Url.focus();
            return false;
    }
    form2.SaveMdyButton.disabled=true
    return true;
}
</script></td>
      <td><input name="Submit4" type="submit" id="SaveMdyButton" class="button01-out" value="ȷ  ��"> 
        <input name="Submit22" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit32" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
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
              <font color="#0000FF">����ɾ��</font></label> <input name="Submit322" type="button" class="button01-out" value="ɾ  ��" onclick="DelReco('<%=Rs("Id")%>',form2.RealDel.checked,'��ȷ��ɾ����')"></td>
          </tr>
        </table>
        <%Else%> 
        <table width="26%" border="0" cellspacing="1" cellpadding="0">
          <tr> 
            <td align="right"><input name="Submit3223" type="button" class="button01-out" value="��  ��" onclick="DelReco('<%=Rs("Id")%>','0','��ȷ��[�ָ�]�˼�¼��')" Title="[�ָ�]��ǰ�Ѿ���[����ɾ��]�ļ�¼"> 
              <input name="Submit3222" type="button" class="button02-out" value="����ɾ��" onclick="DelReco('<%=Rs("Id")%>','1','��ȷ��[����ɾ��]��\n[����ɾ��]�ļ�¼���޷���[�ָ�]')" title="[����ɾ��]�ļ�¼���޷���ԭ"></td>
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

'������Ŀ��
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

'������Ŀ��
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