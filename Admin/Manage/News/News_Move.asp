<!--#include file="../Include/Conn.asp" -->
<!-- #include file="../Include/ClassList_Fun.asp" -->
<!--#include file="../Include/Config.asp" -->
<!--#include file="../Include/Tkl_StringClass.asp" -->
<!--#include file="../Include/Tkl_SYSProedomClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("../Login.asp")
End If

If Not SysAdmin.MoveNews Then
    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End If
%>
<html>
<head>
<title>News_Move</title>
<meta name="Generator" content="EditPlus">
<meta name="Author" content="">
<meta name="Keywords" content="">
<meta name="Description" content="">
<link href="../Include/ManageStyle.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">

<body bgcolor="#FFFFFF">
<%
Dim Work
    Work=Request("Work")
    Select Case Work
        Case "MoveByClass"
            Call MoveByClass()
        Case "MoveById"
            Call MoveById()
        Case "MoveByTime"
            Call MoveByTime()
    End Select
Dim I
%>
<form name="form1" method="post" action="?">
  <script language="JavaScript" src="../Include/Tkl_ClassTree.js" type="text/JavaScript"></script>
  <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
          <tr> 
            <td bgcolor="#FFCC33" id="tab1"><label class="BarTitle" for="work1"> 
              <input name="Work" id="work1" type="radio" value="MoveByClass" checked onClick="changeTab(window.tab1)">
              �ƶ��������</label></td>
          </tr>
          <tr> 
            <td bgcolor="#F6f6f6"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
                <tr bgcolor="#FFFFFF"> 
                  <td width="23%" valign="top" class="BarTitle">���ƶ����:</td>
                  <td width="77%" align="center" valign="middle" class="BarText"><label style="cursor:hand"> 
                    <script>
            var root1
            root1=CreateRoot("myTree1","�����ƶ���Դ���")
            <%Call CreateClassTree1(0)%>
        </script>
                    </label></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td valign="top" class="BarTitle">�ƶ���:</td>
                  <td align="center" valign="middle" class="BarText"><label style="cursor:hand"><font color="#0000FF"><font color="#0000FF"></font></font></label>
                    <script>
                var root2
                root2=CreateRoot("myTree2","��Ŀ�ĵ����")
                <%Call CreateClassTree2(0)%>
            </script> </td>
                </tr>
              </table></td>
          </tr>
        </table>
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="10"></td>
    </tr>
  </table>
  <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
          <tr> 
            <td bgcolor="#FFFFCC" id="tab2"><label class="BarTitle" for="work2"> 
              <input type="radio" id="work2" name="Work" value="MoveById" onClick="changeTab(window.tab2)">
              ָ����¼ID��Χ�ƶ�</label></td>
          </tr>
          <tr> 
            <td bgcolor="#F6f6f6"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
                <tr> 
                  <td width="23%" bgcolor="#FFFFFF" class="BarTitle">���ƶ���¼:</td>
                  <td width="77%" bgcolor="#FFFFFF" class="BarText">��ʼ��Դ��¼ID: 
                    <input name="StartId" type="text" class="Input" id="startId3" value="0" size="6">
                    ������Դ��¼ID: 
                    <input name="EndId" type="text" class="Input" id="EndId" value="0" size="6"></td>
                </tr>
                <tr> 
                  <td valign="top" bgcolor="#FFFFFF" class="BarTitle"> �ƶ���:</td>
                  <td bgcolor="#FFFFFF"> 
                    <script>
                var root3
                root3=CreateRoot("myTree3","��Ŀ�ĵ����")
                <%Call CreateClassTree3(0)%>
            </script></td>
                </tr>
              </table></td>
          </tr>
        </table>		
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="10"></td>
    </tr>
  </table>
  <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
          <tr> 
            <td bgcolor="#FFFFCC" id="tab3"><label class="BarTitle" for="work3"> 
              <input type="radio" id="work3" name="Work" value="MoveByTime" onClick="changeTab(window.tab3)">
              ָ��ʱ�䷶Χ�ƶ�</label></td>
          </tr>
          <tr> 
            <td bgcolor="#F6f6f6"><table width="100%" border="0" cellpadding="2" cellspacing="0">
                <tr id="tr1_1"> 
                  <td width="23%" height="7" bgcolor="#f6f6f6" class="BarText">����:</td>
                  <td width="77%" bgcolor="#FFFFFF"> <select name="TimeType" class="Input" id="select2">
                      <option value="upTime" selected>������ʱ��</option>
                      <option value="AddTime">��Դ����ʱ��</option>
                    </select></td>
                </tr>
                <tr id="tr1_2"> 
                  <td height="8" bgcolor="#f6f6f6" class="BarText">��ʼʱ��:</td>
                  <td bgcolor="#FFFFFF"> <select name="startYear" class="Input">
                      <%For I=Year(Now) To 1900 Step -1%>
                      <option value="<%=I%>"><%=I%></option>
                      <%Next%>
                    </select>
                    - 
                    <select name="startMonth" class="Input">
                      <%For I=1 To 12%>
                      <option value="<%=I%>" <%If I=Month(Now) Then Response.Write("selected") End If%>><%=I%></option>
                      <%Next%>
                    </select>
                    - 
                    <select name="startDay" class="Input">
                      <%For I=1 To 31%>
                      <option value="<%=I%>" <%If I=Day(Now) Then Response.Write("selected") End If%>><%=I%></option>
                      <%Next%>
                    </select></td>
                </tr>
                <tr id="tr1_3"> 
                  <td height="16" bgcolor="#f6f6f6" class="BarText">����ʱ��:</td>
                  <td bgcolor="#FFFFFF"> <select name="EndYear" class="Input">
                      <%For I=Year(Now) To 1900 Step -1%>
                      <option value="<%=I%>"><%=I%></option>
                      <%Next%>
                    </select>
                    - 
                    <select name="EndMonth" class="Input">
                      <%For I=1 To 12%>
                      <option value="<%=I%>" <%If I=Month(Now) Then Response.Write("selected") End If%>><%=I%></option>
                      <%Next%>
                    </select>
                    - 
                    <select name="EndDay" class="Input">
                      <%For I=1 To 31%>
                      <option value="<%=I%>" <%If I=Day(Now) Then Response.Write("selected") End If%>><%=I%></option>
                      <%Next%>
                    </select></td>
                </tr>
                <tr id="tr1_3">
                  <td height="16" valign="top" bgcolor="#f6f6f6" class="BarText">Ŀ�����:</td>
                  <td bgcolor="#FFFFFF" class="BarText"> 
                    <script>
                var root4
                root4=CreateRoot("myTree4","��Ŀ�ĵ����")
                <%Call CreateClassTree4(0)%>
            </script></td>
                </tr>
              </table></td>
          </tr>
        </table>
  <table width="100%" border="0" cellpadding="2" cellspacing="1">
    <tr> 
      <td align="center"><label class="BarTitle">
        <input name="Submit" type="submit" class="button01-out" value="ȷ  ��">
        <input name="Submit2" type="reset" class="button01-out" value="��  ԭ">
        <input name="Submit3" type="button" class="button01-out" value="��  ��" onclick="window.history.back();">
        </label></td>
    </tr>
  </table>
</form>
<script>
var activeTab=window.tab1
function changeTab(obj)
{
    activeTab.bgColor="#FFFFCC"
    obj.bgColor="#FFCC33"
    activeTab=obj
}
</script>
<script language="JavaScript" type="text/JavaScript">
function confimSub()
{
    if(confirm("��ȷ��ִ�е�ǰ���ƶ�������\n�������������")){
        form1.Submit2.click()
        return true
    }else{
        return false
    }
}
</script>
</body>
</html>
<%
Sub CreateClassTree1(ParentId)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        If Rs("Parent")=0 Then
            Response.Write "root1.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""checkbox\"" NAME=\""sourceClass\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root1.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""checkbox\"" NAME=\""sourceClass\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        CreateClassTree1 Rs("Id")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

Sub CreateClassTree2(ParentId)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        If Rs("Parent")=0 Then
            Response.Write "root1.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""radio\"" NAME=\""targetClass\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root1.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""radio\"" NAME=\""targetClass\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        CreateClassTree2 Rs("Id")
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
            Response.Write "root3.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""radio\"" NAME=\""targetClass1\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root3.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""radio\"" NAME=\""targetClass1\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        CreateClassTree3 Rs("Id")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

Sub CreateClassTree4(ParentId)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        If Rs("Parent")=0 Then
            Response.Write "root4.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""radio\"" NAME=\""targetClass2\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root4.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""radio\"" NAME=\""targetClass2\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        CreateClassTree4 Rs("Id")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub

Sub MoveByClass()
    'On Error Resume Next
    Dim sourceClass
        sourceClass=Request("sourceClass")
    Dim targetClass
        targetClass=Request("targetClass")
    If sourceClass="" Or targetClass="" Then
        Response.Write("<script>alert(""<'�ƶ��������'����ʧ��>\n��ѡ��[Դ���]��[Ŀ�����]"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End
    End If
    Dim arrSourClass
        arrSourClass=Split(sourceClass,",",-1,1)
    Dim Sql,I
    For I=0 To UBound(arrSourClass)
        Sql="UPDATE view_NewsInfo SET Class="&targetClass&" WHERE Class="&CLng(Trim(arrSourClass(I)))
        Conn.ExeCute(Sql)
    Next
    Response.Write("<script>alert(""<'�ƶ��������'�����ɹ�>\n��Դ�ƶ����"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End
End Sub

Sub MoveById()
    'On Error Resume Next
    Dim targetClass
        targetClass=Request("targetClass1")
    Dim StartId
        StartId=Request("StartId")
    Dim EndId
        EndId=Request("EndId")
    If targetClass="" Then
        Response.Write("<script>alert(""<'ָ����¼ID��Χ�ƶ�'����ʧ��>\n��ѡ��[Ŀ�����]"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End
    End If
    If Not IsNumeric(StartId) Or Not IsNumeric(EndId) Then
        Response.Write("<script>alert(""<'ָ����¼ID��Χ�ƶ�'����ʧ��>\n����ȷ������Դ��¼��[��ʼID]��[����ID]"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End
    End If
    Dim Sql,I
        Sql="UPDATE view_NewsInfo SET Class="&targetClass&" WHERE "&CLng(StartId)&"<=Id And Id<="&CLng(EndId)
        Conn.ExeCute(Sql)
    Response.Write("<script>alert(""<'ָ����¼ID��Χ�ƶ�'�����ɹ�>\n��Դ�ƶ����"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End
End Sub

Sub MoveByTime()
    'On Error Resume Next
    Dim targetClass
        targetClass=Request("targetClass2")
    If targetClass="" Then
        Response.Write("<script>alert(""<'ָ��ʱ�䷶Χ�ƶ�'����ʧ��>\n��ѡ��[Ŀ�����]"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End
    End If
    Dim startTime,endTime
    startTime=Request("startMonth") &"/" & Request("startDay") & "/" & Request("startYear")     
    endTime=Request("endMonth") &"/" & Request("endDay") & "/" & Request("endYear")
    Dim TimeType
        TimeType=Request("TimeType")
    Dim Sql,I
        Sql="UPDATE view_NewsInfo SET Class="&targetClass&" Where #"&startTime&"#<="&TimeType&" And "&TimeType&"<=#"&endTime&"#"
    Conn.ExeCute(Sql)
    Response.Write("<script>alert(""<'ָ��ʱ�䷶Χ�ƶ�'�����ɹ�>\n��Դ�ƶ����"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End
End Sub
%>