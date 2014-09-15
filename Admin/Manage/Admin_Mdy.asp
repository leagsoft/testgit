<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Tkl_StringClass.asp" -->
<!--#include file="Include/CfsEnCode.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not SysAdmin.Logined Then
    Response.Redirect("Login.asp")
End If

Dim CFun
set CFun=New Tkl_StringClass
%>
<html>
<head>
<title>Admin_Mdy.asp</title>
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
    Set Rs=Conn.ExeCute("Select * From Admin Where Id=" & Request("Id"))
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("��¼δ�ҵ�")
        Response.End
    End If
    Dim Id,Title,Role,mLock,NickName
    Id=Rs("Id")
    Title=Rs("Title")
    NickName=Rs("NickName")
    Role=Rs("Role")
    mLock=Rs("Lock")
    Rs.Close
    Set Rs=Nothing
%>
<form name="form1" method="post" action="?Work=SaveMdy" onSubmit="return checkMdyReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">�༭�ʻ�</td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">ID:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Id" type="text" class="Input" id="Id2"  value="<%=Id%>" size="4" readonly="true"></td>
    </tr>
    <tr> 
      <td class="BarTitle">�ʻ�:</td>
      <td bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title"  value="<%=Title%>" size="40" readonly="true"></td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">�ǳƣ�������:</td>
      <td width="75%" bgcolor="#FFFFFF"><input name="NickName" type="text" class="Input" id="Title5"  value="<%=NickName%>" size="40"> 
      </td>
    </tr>
    <%If SysAdmin.ChangeAdminList And Title<>SysAdmin.defAdminUserTitle Then%>
    <tr> 
      <td class="BarTitle">��ɫ:</td>
      <td width="75%" bgcolor="#FFFFFF"><select name="Role" id="Role">
          <option value="">��ѡ��</option>
          <%
          Dim Rs2
          Set Rs2=Conn.ExeCute("Select Id,Title From Admin_Role Order By ID")
          While Not Rs2.Eof
          %>
          <option value="<%=Rs2("Id")%>" <%If Rs2("Id")=Role Then Response.Write "Selected" End If%>><%=Rs2("Title")%></option>
          <%
              Rs2.MoveNext
          Wend
          Rs2.Close
          %>
        </select></td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">����:</td>
      <td bgcolor="#FFFFFF" ><img src="Images/Skin/Lock.gif" width="16" height="16"><label for="Lock1"><input type="radio" id="Lock1" name="Lock" value="1" <%If CBool(mLock) Then Response.Write "checked" End If%>>
        ����</label>��<img src="Images/Skin/UnLock.gif" width="16" height="16"><label for="UnLock1"><input id="UnLock1" name="Lock" type="radio" value="0" <%If Not CBool(mLock) Then Response.Write "checked" End If%>>����</label></td>
    </tr>
    <%End If%>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkMdyReco(obj){
    if(obj.Title.value==""){
        alert("������[�ʻ�]");
        obj.Title.focus();
        return false;
    }
    if(obj.NickName.value==""){
        alert("������[�ǳ�(����)]");
            obj.NickName.focus();
        return false;
    }
    try{
        if(obj.Role.value==""){
            alert("������[��ɫ]");
                obj.Role.focus();
            return false;
        }
    }catch(exception)
    {}
    return true;
}
</script> </td>
      <td bgcolor="#FFFFFF"> 
        <input name="Submit" type="submit" class="button01-out" value="ȷ  ��">
        <input name="Submit2" type="reset" class="button01-out" value="��  ԭ">
        <input name="Submit3" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
    </tr>
    <tr> 
      <td colspan="2" align="right" bgcolor="#FFFFFF">
<script>
function DelReco(id){
    if(confirm("��ȷ��ɾ����")){
        window.location="?Work=DelReco&Id="+id;
    }
}

function ChangePwd(id){
    var pwd=showModalDialog("Admin_ChangePwd.htm?",id,"dialogWidth:300px;dialogHeight:200px;center:yes;scroll:no;");
}
</script>
        <input name="Submit4" type="button" class="button01-out" onclick="ChangePwd('<%=Id%>')" value="�� ��" <%If Not SysAdmin.ChagePWD Then Response.Write("disabled=""true""") End If%>>
        <input name="Submit5" type="button" class="button01-out" onclick="DelReco('<%=Id%>')" value="ɾ ��" <%If Title=SysAdmin.defAdminUserTitle Or Not SysAdmin.ChangeAdminList Then Response.Write("disabled=""true""") End If%>>
      </td>
    </tr>
  </table>
</form>
<%End Sub%>
<%
Sub AddReco()
%>
<form name="form2" method="post" action="?Work=SaveAddReco" onSubmit="return checkAddReco(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td colspan="2" align="center" class="BarTitleBg">����ʻ�</td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">�ʻ�:</td>
      <td width="75%" bgcolor="#FFFFFF"> <input name="Title" type="text" class="Input" id="Title2" size="40" ></td>
    </tr>
    <tr> 
      <td class="BarTitle">�ǳƣ�������:</td>
      <td bgcolor="#FFFFFF"> <input name="NickName" type="text" class="Input" id="Title5" size="40"> 
      </td>
    </tr>
    <tr> 
      <td class="BarTitle">����:</td>
      <td width="75%" bgcolor="#FFFFFF"><input name="Pwd" type="password" class="Input" id="Pwd" size="40" ></td>
    </tr>
    <tr> 
      <td class="BarTitle">ȷ������:</td>
      <td width="75%" bgcolor="#FFFFFF"><input name="Pwd2" type="password" class="Input" id="Title4" size="40" ></td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">��ɫ:</td>
      <td bgcolor="#FFFFFF"><select name="Role" id="Role">
          <option value="">��ѡ��</option>
<%
          Dim Rs2
          Set Rs2=Conn.ExeCute("Select Id,Title From Admin_Role Order By ID")
          While Not Rs2.Eof
%>
          <option value="<%=Rs2("Id")%>"><%=Rs2("Title")%></option>
<%
              Rs2.MoveNext
          Wend
          Rs2.Close
%>
        </select></td>
    </tr>
    <tr> 
      <td width="25%" class="BarTitle">����:</td>
      <td bgcolor="#FFFFFF"><img src="Images/Skin/Lock.gif" width="16" height="16"><label for="Lock2"><input id="Lock2" name="Lock" type="radio" value="1" checked>����</label>��<img src="Images/Skin/UnLock.gif" width="16" height="16"> 
        <label for="UnLock2"><input id="UnLock2" name="Lock" type="radio" value="0">����</label></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#FFFFFF"> 
        <script>
function checkAddReco(obj){
    if(obj.Title.value==""){
        alert("������[�ʻ�]");
        obj.Title.focus();
        return false;
    }
    if(obj.NickName.value==""){
        alert("������[�ǳ�(����)]");
            obj.NickName.focus();
        return false;
    }    
    if(obj.Pwd.value=="")
    {
        alert("������[����]");
        obj.Pwd.focus();
        return false;    
    }
    if(obj.Pwd2.value=="")
    {
        alert("������[ȷ������]");
        obj.Pwd2.focus();
        return false;    
    }
    if(obj.Pwd2.value!=obj.Pwd.value)
    {
        alert("[����]��[ȷ������]��һ��");
        obj.Pwd.focus();
        return false;
    }
    if(obj.Role.value==""){
        alert("������[��ɫ]");
        obj.Role.focus();
        return false;
    }
    return true;
}
</script> </td>
      <td bgcolor="#FFFFFF"> <input name="Submit6" type="submit" class="button01-out" value="ȷ  ��"> 
        <input name="Submit22" type="reset" class="button01-out" value="��  ԭ"> 
        <input name="Submit32" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
    </tr>
  </table>
</form>
<%End Sub%>
</body>
</html>
<%
Sub SaveMdy()
    If Not SysAdmin.ChangeAdminList Then
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Id,Title,NickName,Role,Pwd,mLock
    Id=Request("Id")
    Role=Request("Role")
    Title=Trim(Request("Title"))
    NickName=Trim(Request("NickName"))
    mLock=Request("Lock")
    Dim Sql        
        If SysAdmin.ChangeAdminList Then
            '�����ʻ�Sql��
            Sql="Select * From Admin Where Id=" & Id
        Else
            '��ͨ�û�Sql��
            '***************************************** Modify By BennyLiu:20040311******************************************************************************
            '**Sql="Select * From Admin Where UCase(Title)='" & UCase(SysAdmin.AdminTitle) & "' And UCase(Title)<>'" & UCase(SysAdmin.defAdminUserTitle) & "'"
            Sql="Select * From Admin Where Title='" & UCase(SysAdmin.AdminTitle) & "' And Title<>'" & UCase(SysAdmin.defAdminUserTitle) & "'"
            '************************************************ End Modify ****************************************************************************************
        End If
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3    
    If Not (Rs.Eof And Rs.Bof) Then
        'AdminUserTitle��Role,Lock���Բ��ܸ���
        If SysAdmin.ChangeAdminList And UCase(Rs("Title"))<>UCase(SysAdmin.defAdminUserTitle) Then
            Rs("Role")= Role
            If CBool(mLock) Then
                Rs("Lock")= 1
            Else
                Rs("Lock")= 0
            End If
        End If
        Rs("NickName")= NickName
        Rs("upTime")= Now
        Rs.Update
    End If
    Rs.Close
    Set Rs=Nothing    
    Response.Redirect("Admin_List.asp")
End Sub

Sub DelReco()
    If Not SysAdmin.ChangeAdminList Then
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If
    
    Dim Sql
    'Ĭ�ϵĳ����ʻ�(SysAdmin.defAdminUserTitle)�޷���ɾ��
    '******************************************************** Modify By BennyLiu:20040311****************************************************
    '**Sql="Delete From Admin Where Id=" & CLng(Request("Id")) & " And UCase(Title)<>'" & UCase(SysAdmin.defAdminUserTitle) & "'"
    Sql="Delete From Admin Where Id=" & CLng(Request("Id")) & " And Title<>'" & UCase(SysAdmin.defAdminUserTitle) & "'"
    '************************************************************ End Modify ****************************************************************
    Conn.ExeCute(Sql)

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "ִ�й���Ա(Id:"&Request("Id")&")ɾ��")
    Set LogClass=Nothing

    Response.Redirect("Admin_List.asp")
End Sub

Sub SaveAddReco()
    If Not SysAdmin.ChangeAdminList Then
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If
    
    Dim Title,NickName,Role,Pwd,mLock
    Role=Request("Role")
    Title=Trim(Request("Title"))
    NickName=Trim(Request("NickName"))
    Pwd=Request("Pwd")
    mLock=Request("Lock")
    If Not CFun.IsChar26AndInt(Title) Then
        Response.Write("<script>alert(""<����ʧ��>\n[�ʻ�����]ֻ������26��Ӣ����ĸ�����ֵ����(���»���)"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If
    Dim Sql
    '**********************Modify By BennyLiu:20040311*********************************************
        'Sql="Select Top 1 * From Admin Where UCase(Title)='"& UCase(Title) &"' Order By ID DESC"
        Sql="Select Top 1 * From Admin Where Title='"& UCase(Title) &"' Order By ID DESC"
    '*******************************End Modify*****************************************************
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
        Rs.Open Sql,Conn,1,3
    If Not(Rs.Eof And Rs.Bof) Then                        
        Response.Write("<script>alert(""<����ʧ��>\n�Ѵ�����ͬ���ʻ�"& SoftCopyright_Script &""");window.history.back();</script>")
        Rs.Close
        Set Rs=Nothing
        Response.End()
    End If    
    Rs.AddNew
    Rs("Title")= Title
    Rs("NickName")=NickName
    Rs("Role")= Role
    Rs("Pwd")= CfsEnCode(Pwd)
    Rs("Lock")=mLock
    Rs("upTime")= Now
    Rs("AddTime")= Now
    Rs.Update
    Rs.Close
    Set Rs=Nothing

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "��ӹ���Ա("&Title&")")
    Set LogClass=Nothing

    Response.Redirect("Admin_List.asp")
End Sub

Sub CreateClassTree2(ParentId,mClassPodomeList)
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
        CreateClassTree1 Rs("Id"),mClassPodomeList
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub
%>