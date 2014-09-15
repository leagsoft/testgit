<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#include file="Include/Tkl_StringClass.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not SysAdmin.Logined Then
    Response.Redirect("Login.asp")
End If

Dim CFun
Set CFun=New Tkl_StringClass
%>
<html>
<head>
<title>AdminRole_Mdy.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>
<body bgcolor="#FFFFFF">
<script language="JavaScript" src="Include/Tkl_ClassTree.js" type="text/JavaScript"></script>
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
<%
Sub MdyReco()
    Dim Rs
    Set Rs=Conn.ExeCute("Select * From Admin_Role Where Id=" & Request("Id"))
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("��¼δ�ҵ�")
        Response.End
    End If
    Dim Id,Title,Content,Popedom,ClassPopedomList,ClassId
    Id=Rs("Id")
    Title=Rs("Title")
    Content=Rs("Content")
    Popedom=Rs("Popedom")
    ClassPopedomList=Rs("ClassPopedom")
	ClassId=Rs("ClassId")
    Rs.Close
    Set Rs=Nothing
%>

<form name="form1" method="post" action="?Work=SaveMdy" onSubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">�༭��ɫ</td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">ID:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Id" type="text" class="Input" id="Id2"  value="<%=Id%>" size="4" readonly="true"></td>
    </tr>
    <tr> 
      <td class="BarTitle">��ɫ����:</td>
      <td bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title"  value="<%=Title%>" size="40" <%If UCase(Title)=UCase(SysAdmin.defAdminRoleTitle) Then Response.Write("readonly=""true""") End If%>></td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">���:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Content" type="text" class="Input" id="Content" value="<%=Content%>" size="40"></td>
    </tr>
    <tr> 
      <td width="25%" valign="top" class="BarTitle">��ɫȨ��:<br>
        <font color="#999999">(ʹ��Ctrl\Shift��ϼ����ж���ѡ��) </font></td>
      <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr> 
            <td><select name="Popedom" size="10" multiple id="Popedom" <%If UCase(Title)=UCase(SysAdmin.defAdminRoleTitle) Then Response.Write("Disabled=""true""") End If%>>
              <%
            Dim I
            Dim PList
            PList=Split(SysAdmin.defPopedomList,",",-1,1)
            For I=0 To UBound(PList)-1
%>
              <option value="<%=PList(I)%>" <%If CFun.ItemInList(Popedom,PList(I)) Then Response.Write("Selected") End If%>><%=Left(PList(I+1)&"����������������������",50)%></option>
              <%
              I=I+1
            Next
%>
            </select> </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td valign="top" class="BarTitle">�������Ȩ������:</td>
      <td bgcolor="#FFFFFF">
      <script>
            var root1
            root1=CreateRoot("myTree1","�������б�")
            <%Call CreateClassTree2 (0,ClassPopedomList&"")%>
      </script>
      </td>
    </tr>
    <tr>
      <td valign="top" class="BarTitle">���ƽ��ɲ鿴�ķ��ࣺ</td>
      <td bgcolor="#FFFFFF">
        <script>
		var root4
		root4=CreateRoot("myTree3","�������б�")
		<%If ClassId=0 Then%>
		root4.CreateNode(0,-1,"<input type=\"radio\" name=\"ClassId\" value=\"0\" checked>������Ŀ")
		<%Else%>
		root4.CreateNode(0,-1,"<input type=\"radio\" name=\"ClassId\" value=\"0\">������Ŀ")		
		<%End If%>
		<%Call CreateClassTree4 (0,ClassId)	%>
	</script>
      </td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkMdyReco(obj){
    if(obj.Title.value==""){
        alert("������[����]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("������[���]");
        obj.Content.focus();
        return false;
    }
    return true;
}
</script> </td>
      <td bgcolor="#FFFFFF"> <input name="Submit" type="submit" class="button01-out" value="ȷ  ��"> 
        <input name="Submit2" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit3" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
    </tr>
    <tr> 
      <td colspan="2" align="right" bgcolor="#FFFFFF"> <script>
function DelReco(id){
    if(confirm("��ȷ��ɾ����")){
        window.location="?Work=DelReco&Id="+id;
    }
}
</script>
        <input name="Submit5" type="button" class="button01-out" onclick="DelReco('<%=Id%>')" value="ɾ ��" <%If Title=SysAdmin.defAdminRoleTitle Or Not SysAdmin.ChangeAdminList Then Response.Write("disabled") End If%>>
        </td>
    </tr>
  </table>
</form>
<%End Sub%>
<%Sub AddReco()%>
<form name="form2" method="post" action="?Work=SaveAddReco" onSubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">������ɫ</td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">��ɫ����:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title2" size="40" ></td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">���:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Content" type="text" class="Input" id="Content3" size="40" ></td>
    </tr>
    <tr> 
      <td valign="top" class="BarTitle">��ɫȨ��:</td>
      <td width="75%" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr> 
            <td> <select name="Popedom" size="10" multiple id="Popedom">
<%
            Dim I
            Dim PList
            PList=Split(SysAdmin.defPopedomList,",",-1,1) 
            For I=0 To UBound(PList)-1
%>
                <option value="<%=PList(I)%>"><%=Left(PList(I+1)&"����������������������",50)%></option>
<%
              I=I+1
            Next
%>
              </select> </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td valign="top" class="BarTitle">�������Ȩ������:</td>
      <td bgcolor="#FFFFFF"> <script>
            var root1
            root1=CreateRoot("myTree1","�������б�")
            <%Call CreateClassTree1(0)%>
        </script></td>
    </tr>
    <tr>
      <td valign="top" class="BarTitle">���ƽ��ɲ鿴�ķ��ࣺ</td>
      <td bgcolor="#FFFFFF">
        <script>
		var root3
		root3=CreateRoot("myTree3","�������б�")
		root3.CreateNode(0,-1,"<input type=\"radio\" name=\"ClassId\" value=\"0\" checked>������Ŀ")
		<%Call CreateClassTree3 (0)%>
	</script>
      </td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkAddReco(obj){
    if(obj.Title.value==""){
        alert("������[����]");
        obj.Title.focus();
        return false;
    }
    if(obj.Content.value==""){
        alert("������[���]");
        obj.Content.focus();
        return false;
    }
    return true;    
}
</script> </td>
      <td bgcolor="#FFFFFF"> <input name="Submit4" type="submit" class="button01-out" value="ȷ  ��"> 
        <input name="Submit22" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit32" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
    </tr>
  </table>
</form>
<%End Sub%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display='';window.scrollTo(window.pageXOffset,2000);}else{HelpTab.style.display='none'}">&nbsp;<img src="Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:none"> 
    <td width="2%">&nbsp;</td>
    <td width="98%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td>���������Ȩ������</td>
        </tr>
        <tr> 
          <td> <p>����Ȩֵ˵����</p></td>
        </tr>
        <tr> 
          <td> ������&quot;<font color="#0000FF">��</font>&quot;:���,�޸����ѵ���Դ��<br>
            ������&quot;<font color="#0000FF">��</font>&quot;:���,�޸������˵���Դ��<br>
            ������&quot;<font color="#FF0000">��</font>&quot;:���,�޸�,���,����,ɾ�������˵���Դ��</td>
        </tr>
        <tr> 
          <td>�������ʻ���ǰ��ɫ�ģ���ɫȨ�ޣ��а�����&quot;<font color="#0000FF">����������Դ</font>&quot;,��'�������Ȩ��'��������,��������з����һ�в���Ȩ��</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <a name="Help"></a></td>
  </tr>
</table>
</body>
</html>
<%
Sub SaveMdy()
	If Not SysAdmin.ChangeRole Then
		Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
		Response.End()
	End If

    Dim Id,Title,Content,Popedom,ClassId
    Id=Request("Id")
    Title=Replace(Trim(Request("Title")),"'","''")
    Content=Replace(Trim(Request("Content")),"'","''")
    Popedom=Trim(Request("Popedom"))
	ClassId=CLng(Request("ClassId"))	

    Dim Rs
    Dim Sql
        Sql="Select * From Admin_Role Where Id=" & Id
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    If Not (Rs.Eof And Rs.Bof) Then
        'AdminRoleTitle��ɫ�Ĳ�����Ϣ���ɱ�����
        If Not UCase(Rs("Title"))=UCase(SysAdmin.defAdminRoleTitle) Then
            '�Ƿ����ͬ����ɫ
            '********************************* Modify By BennyLiu:20040311************************************
            'Sql="Select * From Admin_Role Where Id<>" & Id & " And UCase(Title)='" & UCase(Title) &"'"
            Sql="Select * From Admin_Role Where Id<>" & Id & " And Title='" & UCase(Title) &"'"
            
            '********************************** End Mofify ****************************************************
            Dim Rs2
            Set Rs2=Conn.ExeCute(Sql)
            If Not(Rs2.Eof And Rs2.Bof) Then
                Response.Write("<script>alert(""<����ʧ��>\nΪ������,���������ͬ��[��ɫ]"& SoftCopyright_Script &""");window.history.back();</script>")        
                Rs.Close
                Set Rs=Nothing
                Rs2.Close
                Set Rs2=Nothing    
                Response.End
            Else
                Rs("Title")= Title
            End If
            Rs2.Close
            Set Rs2=Nothing
            Rs("Popedom")= Popedom
        End If
        Rs("ClassPopedom")=BalePopedomChar()
        Rs("Content")= Content
		Rs("ClassId")=ClassId		
        Rs("upTime")= Now
        Rs.Update
    End If
    Rs.Close
    Set Rs=Nothing

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "�޸Ľ�ɫ("&Title&")")
    Set LogClass=Nothing

    Response.Redirect("AdminRole_List.asp")
End Sub

Sub DelReco()
    If Not SysAdmin.ChangeRole Then
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Sql
    Sql="Select Count(Id) As Num From Admin Where Role=" & Request("Id")
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    If Rs("Num")>=1 Then
        Response.Write("<script>alert(""<����ʧ��>\n��ǰ���������û�����ʹ�ô�[��ɫ],����޷�ɾ��!"& SoftCopyright_Script &""");window.history.back();</script>")        
        Rs.Close
        Set Rs=Nothing
        Response.End()
    End If
    
    '******************************Modify By BennyLiu:20040311**********************************************************
    '**Sql="Delete From Admin_Role Where Id=" & Request("Id") &" And UCase(Title)<>'"& UCase(SysAdmin.defAdminRoleTitle) &"'"
    Sql="Delete From Admin_Role Where Id=" & Request("Id") &" And Title<>'"& UCase(SysAdmin.defAdminRoleTitle) &"'"
    '***************************************** End Modify********************************************************
    Conn.ExeCute(Sql)

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "ɾ����ɫ(Id:"&Request("Id")&")")
    Set LogClass=Nothing

    Response.Redirect("AdminRole_List.asp")
End Sub

Sub SaveAddReco()
	If Not SysAdmin.ChangeRole Then
		Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
		Response.End()
	End If

    Dim Title,Content,Popedom,ClassId
    Title=Replace(Trim(Request("Title")),"'","''")
    Content=Replace(Trim(Request("Content")),"'","''")
    Popedom=Trim(Request("Popedom"))
	ClassId=CLng(Request("ClassId"))

    Dim Sql
        Sql="Select Top 1 * From Admin_Role Where Title='" & Title & "' Order By ID DESC"
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    If Not(Rs.Eof And Rs.Bof) Then
        Response.Write("<script>alert(""<����ʧ��>\nΪ������,���������ͬ��[��ɫ]"& SoftCopyright_Script &""");window.history.back();</script>")        
        Rs.Close
        Set Rs=Nothing
        Response.End()
    End If
    Rs.AddNew
    Rs("Title")= Title
    Rs("Content")= Content
    Rs("Popedom")= Popedom
    Rs("ClassPopedom")=BalePopedomChar()
	Rs("ClassId")=ClassId
    Rs("upTime")= Now
    Rs.Update
    Rs.Close
    Set Rs=Nothing

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "������ɫ("&Title&")")
    Set LogClass=Nothing

    Response.Redirect("AdminRole_List.asp")
End Sub

'//����'�������Ȩ������',������������ݸ�ʽ��
'//����Id1&","&Ȩֵ1& vbCrLf &����Id2&","&Ȩֵ2&vbCrLf&....
'//����:string
Function BalePopedomChar()
    BalePopedomChar=""
    Dim ClassId_List
        ClassId_List=Replace(Trim(Request("sourceClass"))&""," ","")
    If ClassId_List="" Then
        Exit Function
    End If
    Dim arrClassId_Item
        arrClassId_Item=Split(ClassId_List,",",-1,1)
    Dim I
    For I=0 To UBound(arrClassId_Item)
        If BalePopedomChar="" Then
            BalePopedomChar = arrClassId_Item(I) & "," & Request("PopedomType" & arrClassId_Item(I))
        Else
            BalePopedomChar = BalePopedomChar & vbCrLf & arrClassId_Item(I)  & "," & Request("PopedomType"&arrClassId_Item(I))
        End If
    Next
End Function

Sub CreateClassTree1(ParentId)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        If Rs("Parent")=0 Then
            Response.Write "root1.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""checkbox\"" NAME=\""sourceClass\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&"&nbsp;&nbsp;<font color=\""blue\""><INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Low&"\"" checked>��&nbsp;<INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Mid&"\"">��</font>&nbsp;<INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Hig&"\""><font color=\""red\"">��</font>"")" & vbCrLf
        Else
            Response.Write "root1.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""checkbox\"" NAME=\""sourceClass\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&"&nbsp;&nbsp;<font color=\""blue\""><INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Low&"\"" checked>��&nbsp;<INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Mid&"\"">��</font>&nbsp;<INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Hig&"\""><font color=\""red\"">��</font>"")" & vbCrLf
        End If
        CreateClassTree1 Rs("Id")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

Sub CreateClassTree2(ParentId,mClassPopedom)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    Dim arrPopedomList
    Dim I
    Dim arrPopedomItem
    Dim strChecked,strPopedomType_Low,strPopedomType_Mid,strPopedomType_Hig
    While Not Rs.Eof
        strChecked=""
        strPopedomType_Low=""
        strPopedomType_Mid=""
        strPopedomType_Hig=""
        arrPopedomList=Split(mClassPopedom,vbCrLf,-1,1)
        For I=0 To UBound(arrPopedomList)
            'ȡ�����ո���Ȩ��
            arrPopedomItem=Split(arrPopedomList(I),",",-1,1)
            If CLng(arrPopedomItem(0))=Rs("Id") Then
                strChecked="Checked"
                Select Case CLng(arrPopedomItem(1))
                    Case SysAdmin.defClassPopedomType_Low
                        strPopedomType_Low="checked"
                    Case SysAdmin.defClassPopedomType_Mid
                        strPopedomType_Mid="checked"
                    Case SysAdmin.defClassPopedomType_Hig
                        strPopedomType_Hig="checked"
                End Select
                Exit For
            End If
        Next
        If Rs("Parent")=0 Then
            Response.Write "root1.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""checkbox\"" NAME=\""sourceClass\"" value=\"""&Rs("Id")&"\"" "&strChecked&">"&Rs("Title")&"&nbsp;&nbsp;<font color=\""blue\""><INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Low&"\"" "&strPopedomType_Low&">��&nbsp;<INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Mid&"\"" "&strPopedomType_Mid&">��</font>&nbsp;<INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Hig&"\"" "&strPopedomType_Hig&"><font color=\""red\"">��</font>"")" & vbCrLf
        Else
            Response.Write "root1.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""checkbox\"" NAME=\""sourceClass\"" value=\"""&Rs("Id")&"\"" "&strChecked&">"&Rs("Title")&"&nbsp;&nbsp;<font color=\""blue\""><INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Low&"\"" "&strPopedomType_Low&">��&nbsp;<INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Mid&"\"" "&strPopedomType_Mid&">��</font>&nbsp;<INPUT TYPE=\""radio\"" NAME=\""PopedomType"&Rs("Id")&"\"" value=\"""&SysAdmin.defClassPopedomType_Hig&"\"" "&strPopedomType_Hig&"><font color=\""red\"">��</font>"")" & vbCrLf
        End If
        CreateClassTree2 Rs("Id"),mClassPopedom
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

Sub CreateClassTree3(ParentId)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        If Rs("Parent")=0 Then
            Response.Write "root3.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""radio\"" NAME=\""ClassId\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root3.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""radio\"" NAME=\""ClassId\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        CreateClassTree3 Rs("Id")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

Sub CreateClassTree4(ParentId,ClassId)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
	Dim strIsChecked
		strIsChecked=""
    While Not Rs.Eof
		If Rs("Id")=ClassId Then
			strIsChecked="checked"
		End If
        If Rs("Parent")=0 Then			
            Response.Write "root4.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""radio\"" NAME=\""ClassId\"" value=\"""&Rs("Id")&"\"" "&strIsChecked&">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root4.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""radio\"" NAME=\""ClassId\"" value=\"""&Rs("Id")&"\"" "&strIsChecked&">"&Rs("Title")&""")" & vbCrLf
        End If
		strIsChecked=""
        CreateClassTree4 Rs("Id"),ClassId
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub
%>