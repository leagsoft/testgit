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
QXMC=Session("QXMC")		'ȡ��Ȩ������
Column=Session("Column")	'�û����Ƶ���Ŀ		'Add Benny

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
      <td colspan="2" class="BarTitleBg"> �����Դ </td>
    </tr>
    <tr> 
      <td width="16%" height="9" valign="top" class="BarTitle">��Դ����:</td>
      <td width="84%" height="9" bgcolor="#FFFFFF"> 
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="68%">
    <%	if Session("YHZL")<>"����Ա" then  %>
    <script language="javascript">
      var root1
      root1=CreateRoot("myTree1","����ѡ��[��Դ���]")
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
<%	elseif Session("YHZL")="����Ա" then%>      
    <script language="javascript">
      var root1
      root1=CreateRoot("myTree1","����ѡ��[��Դ���]")
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
    <tr> 
      <td width="16%" class="BarTitle">��ת����:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="Url" type="text" class="Input" id="Url" size="60">(���ô˹���ʱ���ÿ�) 
      </td>
    </tr>
    <%
    '*******Add By BennyLiu:20040618***********
    '***�ж��Ƿ��Ƿ����ֳ�ר�⣬�ǵĻ�����������ѡ��ֳ���********
    If QXMC="�ֳ�ר��" then
		cSql="select YHDL,UserRankRight from YHXX where YHBM='���쵼' order by UserRankRight"
		set cRs=server.CreateObject ("Adodb.Recordset")
		cRs.Open cSql,connect,1,3    
    %>
    <tr> 
      <td width="16%" class="BarTitle">��Դ����:</td>
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
      <td width="16%" class="BarTitle">��Դ����:</td>
      <td width="84%" bgcolor="#FFFFFF"><input type="text" name="Author" class="Input" id="Author" value="<%=session("YHDL")%>">
      &nbsp;&nbsp;<font color="red">*</font></td>
    </tr>
    <%End If%>
    <tr> 
      <td width="16%" class="BarTitle">��Դ:</td>
      <td width="84%" bgcolor="#FFFFFF"><input type="text" name="From" class="Input" id="From" value="<%=Session("YHBM")%>">&nbsp;&nbsp;<font color="red">*</font></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">���α༭:</td>
      <td bgcolor="#FFFFFF"><input name="Editor" type="text" class="Input" value="<%=Session("YHDL")%>">
      </td>
    </tr>
<!--Add by BennyLiu:20040712-->    
    <%if QXMC="�־ֶ�̬" then%>
    <tr> 
      <td width="16%" class="BarTitle">����ߣ�</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=group&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=user&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Session("Browser")%>">
      </td>
    </tr>
    <%end if%> 
<!--End Add-->
    <%'Add By BennyLiu:20040625   Ϊ�˶������ͳ����Ϣ�������
	if QXMC="����ͳ����Ϣ" then
    %>
    <tr> 
      <td width="16%" class="BarTitle">����ߣ�</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=group&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=user&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form1.Browser&value='+form1.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Session("Browser")%>">
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
    <%end if%>    
    <tr> 
      <td width="16%" valign="top" class="BarTitle">�ؼ���:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="KeyWord" type="text" class="Input" id="Map" size="50" onmouseover="showToolTip('��[�ؼ���]֮����ʹ�á����š��������磺<br><b>����,��Ϸ</b>',event.srcElement)" onmouseout="hiddenToolTip()"></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">���:</td>
      <td bgcolor="#FFFFFF"><input name="Count" type="text" class="Input" id="Count" value="1" size="4"></td>
    </tr>
    <tr>
        <td height=23 colspan="2" class="BarTitle">
		<textarea name="logtext" style="display:none" id="body"></textarea>
		<!--#include file="htmedit.asp"-->
		</td>
	</tr>     
    <!--<tr> 
      <td height="22" colspan="2" valign="top" class="BarTitle"><font color="#0000FF">��Դ����</font>: 
        <span style="cursor:hand" Title="�༭���Ӹ�" onClick="tdNewsContent.style.height=1000;">[�Ӹ�]</span>&nbsp;<span style="cursor:hand" Title="�༭��Ĭ�ϸ߶�" onClick="tdNewsContent.style.height=400;">[Ĭ��]</span>&nbsp;<span style="cursor:hand" Title="�༭������" onClick="tdNewsContent.style.height=200;">[����]</span>&nbsp;<span style="cursor:hand" Id="ShowHiddenHtmlEdit" Title="��ʾ/����" onClick="if(trNewsContent.style.display==''){trNewsContent.style.display='none';ShowHiddenHtmlEdit.innerHTML='[��ʾ]'}else{trNewsContent.style.display='';ShowHiddenHtmlEdit.innerHTML='[����]'}">[����] 
        </span>&nbsp;&nbsp;<font color="red">*</font></td>
    </tr>
    <tr id="trNewsContent"> 
      <td height="400" colspan="2" valign="top" bgcolor="buttonface" Id="tdNewsContent1"><textarea name="NewsContent"></textarea></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">�ϴ�:</td>
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
        alert("��ѡ��[��Դ���]")
        return false
    }
    if(obj.Title.value==""){
        alert("������[��Դ����]");
        obj.Title.focus();
        return false;
    }
    //if(obj.Url.value!=""){
        //if(obj.Url.value.search(/^[a-z0-9]+:\/\/[a-z0-9]+/i)==-1)
       // {
       //     alert("[��ת����]��ʽ����");
      //      obj.Url.focus();
     //       return false;
    //    }
   // }else{
    
    //}
   // if(obj.Author.value==""){
   //     alert("��ѡ��[��Դ����]");
   //     obj.Author.focus();
   //     return false;
   // }
    //if(obj.From.value==""){
     //   alert("��ѡ��[��Դ��Դ]");
    //    obj.From.focus();
    //    return false;
    //}
    //if(obj.KeyWord.value==""){
       // alert("������[��Դ�ؼ���]");
        //obj.KeyWord.focus();
        //return false;
    //}        
    if(obj.Count.value==""){
        alert("������[��Դ�����]");
        obj.Count.focus();
        return false;
    }
    if(obj.Url.value=="")
    {
        if(obj.logtext.value==""){
            alert("������[��Դ����]");
            return false;
        }
    }
    //if(obj.ShortContent.value==""){
    //    alert("������[��ԴժҪ]");
	//	obj.ShortContent.focus();
    //    return false;
    //}
    obj.SaveAddButton.disabled=true
    return true;
}
</script> </td>
      <td><input name="SaveAddButton" type="submit" id="SaveAddButton" class="button01-out" value="ȷ  ��"> 
        <input name="Submit2" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit3" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
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
    <tr> 
      <td width="16%" class="BarTitle">��ת����:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="Url" type="text" class="Input" id="Url" size="60" value="<%=cFun.HTMLEncode2(Rs("Url"))%>">(���ô˹���ʱ���ÿ�) 
      </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">��Դ����:</td>
      <%
      'Modify By BennyLiu:20040625
      If QXMC="�ֳ�ר��" then
			cSql2="select YHDL,UserRankRight from YHXX where YHBM='���쵼' order by UserRankRight"
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
      <td width="16%" class="BarTitle">��Դ:</td>
      <td width="84%" bgcolor="#FFFFFF"><input type="text" name="From" class="Input" id="From" value="<%=session("YHBM")%>">&nbsp;&nbsp;<font color="red">*</font></td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">���α༭:</td>
      <td bgcolor="#FFFFFF"><input name="Editor" type="text" class="Input" value="<%=Session("YHDL")%>">
      </td>
    </tr>
<!--Add by BennyLiu:20040712-->
    <%if QXMC="�־ֶ�̬" then%>
    <tr> 
      <td width="16%" class="BarTitle">�����:</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=group&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=user&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
		<input name="Browser" type="text" class="Input" size="50" value="<%=Rs("Browser")%>">
      </td>
    </tr>   
    <%end if%>
<!--End Add-->    
    <%'Add by BennyLiu:20040625  Ϊ�˶��������ߣ�ֻ�н���ͳ����Ϣ���ܴ��
	if QXMC="����ͳ����Ϣ" then
    %>
    <tr> 
      <td width="16%" class="BarTitle">�����:</td>
      <td bgcolor="#FFFFFF">
		<!--<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=group&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/info.asp?page=30&type=user&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>-->
		<input type=button value="���Ⱥ��" onclick="javascript:window.open('../../SetPurview/infogroup.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button1 name=button1>
		<input type=button value="����û�" onclick="javascript:window.open('../../SetPurview/infouser.asp?p=1&field=form2.Browser&value='+form2.Browser.value,'','Width=760,Height=500,scrollbars=yes');" id=button2 name=button2>
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
    <%end if%>    
    <tr> 
      <td width="16%" valign="top" class="BarTitle">�ؼ���:</td>
      <td width="84%" bgcolor="#FFFFFF"><input name="KeyWord" type="text" class="Input" id="Map" size="50" value="<%=cFun.HTMLEncode2(Rs("KeyWord"))%>">
        (���ؼ��ʼ���[����]����) </td>
    </tr>
    <tr> 
      <td width="16%" class="BarTitle">���:</td>
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
      <td height="19" colspan="2" valign="top" class="BarTitle"><font color="#0000FF">��Դ����</font>: 
        <span style="cursor:hand" Title="�༭���Ӹ�" onClick="tdNewsContent.style.height=1000;">[�Ӹ�]</span>&nbsp;
        <span style="cursor:hand" Title="�༭��Ĭ�ϸ߶�" onClick="tdNewsContent.style.height=400;">[Ĭ��]</span>&nbsp;
        <span style="cursor:hand" Title="�༭������" onClick="tdNewsContent.style.height=200;">[����]</span>&nbsp;
        <span style="cursor:hand" Id="ShowHiddenHtmlEdit" Title="��ʾ/����" onClick="if(trNewsContent.style.display==''){trNewsContent.style.display='none';ShowHiddenHtmlEdit.innerHTML='[��ʾ]'}else{trNewsContent.style.display='';ShowHiddenHtmlEdit.innerHTML='[����]'}">[����]</span>
		&nbsp;&nbsp;<font color="red">*</font> 
      </td>
    </tr>
    <tr id="trNewsContent"> 
      <td height="400" colspan="2" valign="top" bgcolor="buttonface" Id="tdNewsContent"> <textarea name="NewsContent"><%If (Not IsNull(Rs("Content"))) Or (Not ""<>Rs("Content")) Then Response.Write Server.HtmlEnCode(replace(Rs("Content"),"11.36.19.2","10.100.0.2"))End If%></textarea></td>
    </tr>
    <tr> 
      <td class="BarTitle">�ϴ�:</td>
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
        alert("��ѡ��[��Դ���]")
        return false
    }
    if(obj.Title.value==""){
        alert("������[��Դ����]");
        obj.Title.focus();
        return false;
    }
    //if(obj.Url.value!=""){
       // if(obj.Url.value.search(/^[a-z0-9]+:\/\/[a-z0-9]+/i)==-1)
       // {
       //     alert("[��ת����]��ʽ����");
      //      obj.Url.focus();
     //       return false;
     //   }
    //}
    //if(obj.Author.value==""){
      //  alert("��ѡ��[��Դ����]");
      //  obj.Author.focus();
     //   return false;
    //}
    //if(obj.From.value==""){
       // alert("��ѡ��[��Դ��Դ]");
      //  obj.From.focus();
     //   return false;
    //}
    //if(obj.KeyWord.value==""){
      //  alert("������[��Դ�ؼ���]");
       // obj.KeyWord.focus();
       // return false;
   // }        
    if(obj.Count.value==""){
        alert("������[��Դ�����]");
        obj.Count.focus();
        return false;
    }    
    if(obj.Url.value=="")
    {
        if(obj.logtext.value==""){
            alert("������[��Դ����]");
            return false;
        }
    }
    //if(obj.ShortContent.value==""){
    //    alert("������[��ԴժҪ]");
	//	obj.ShortContent.focus();
    //} 
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display='';window.scrollTo(window.pageXOffset,2000);}else{HelpTab.style.display='none'}">&nbsp;<img src="Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:none"> 
    <td width="2%">&nbsp;</td>
    <td width="98%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <!--<tr> 
          <td>������/���¾�̬�ļ�:</td>
        </tr>
        <tr> 
          <td> ����δ��˵���Դ����ִ�д˲���</td>
        </tr>-->
        <tr> 
          <td>��'ɾ��'��'�ָ�'��'����ɾ��'</td>
        </tr>
        <tr> 
          <td> ����1.'ɾ��':������¼���߼����Ϊ'��ɾ��',��ͨ��'����վ'�лָ��ü�¼��'ɾ��'��ͬʱϵͳҲ�����Ӧ�ľ�̬��Դ�ļ�����ɾ��</td>
        </tr>
        <tr> 
          <td>����2.'����ɾ��':���޷���ر�'����ɾ��'����Դ��¼</td>
        </tr>
        <tr> 
          <td>���Զ���ҳ��ǩ:</td>
        </tr>
        <tr> 
          <td>�������ñ�ǩ������Ӧ��HtmlԴ��λ�õ���,ϵͳ���Ը÷�ҳ��ǩ����λ�ý���Դ���ݷָ�ɶ��ҳ���������</td>
        </tr>
        <tr> 
          <td>����ת����:</td>
        </tr>
        <tr> 
          <td>����������Դ�ڱ������ֱ����ת��ָ����[��ת����]��ַ������ʹ�ô˹������ÿա�</td>
        </tr>
        <tr> 
          <td>����Դ������ʽ:</td>
        </tr>
        <tr> 
          <td>������Դ��������ʹ��Html��ǩ����ɫ�ʡ������趨���磺��&lt;font color='red'&gt;�����������ű���&lt;/font&gt;������ע�⣬���е�˫���ű���ĳɵ����š�</td>
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

'������Ŀ��
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

'������Ŀ��
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