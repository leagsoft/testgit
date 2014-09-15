<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If
%>
<html>
<head>
<title>DataBase_Statistic.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link href="Include/ManageStyle.css" rel="stylesheet" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
<script language="JavaScript">
<!--
function SelStatic(num)
{
    var url=""
    switch(num)
    {
        case 1:
            url="?Work=Static01"
            break
        case 2:
            url="?Work=Static02"
            break
        case 3:
            url="?Work=Static03"
            break
        case 4:
            url="?Work=Static04"
            break
        case 5:
            url="?Work=Static05"
            break
    }
    window.location=url
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF">
<%
Dim Work
    Work=Request("Work")
Dim STitle
Select Case Work
    Case "Static01"
        STitle="��Ҫͳ��"
    Case "Static02"
        STitle="��Աͳ��"
    Case "Static03"
        STitle="��Դ�ֲ������"
    Case "Static04"
        STitle="��Դ�ֲ���Сʱ"
    Case "Static05"
        STitle="��Դ�ֲ�����Ŀ"
    Case Else
        STitle="��Ҫͳ��"
End Select
%>
<table width="100%" height="192" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr align="center"> 
    <td height="10" colspan="2" class="BarTitleBg">��<%=STitle%>�����ݿ�ͳ�� </td>
  </tr>
  <tr align="center" bgcolor="#FFFFCC"> 
    <td height="10" colspan="2"> 
      <input name="Submit32" type="button" class="button01-out" value="��Ҫͳ��" onClick="SelStatic(1)">
      <input name="Submit322" type="button" class="button01-out" value="��Աͳ��" onClick="SelStatic(2)" title="ϵͳ����������ϵͳ�ʻ��ڸ���ɫ�еķֲ����">
      <input name="Submit323" type="button" class="button02-out" value="�·���ӷֲ�" onClick="SelStatic(3)" title="ϵͳ������ָ������и��·���Դ����ӷֲ����">
      <input name="Submit3232" type="button" class="button02-out" value="Сʱ��ӷֲ�" onClick="SelStatic(4)" title="ϵͳ������ָ������и�Сʱ��Դ����ӷֲ����">
      <input name="Submit32322" type="button" class="button02-out" value="��Ŀ��Դ�ֲ�" onClick="SelStatic(5)" title="ϵͳ������ָ���������Ŀ��Դ����ӷֲ����">
    </td>
  </tr>
  <tr> 
    <td width="19%" height="168" align="center" valign="top" bgcolor="#FFFFFF"><img src="Images/Manage/Statistic.gif" width="136" height="125"></td>
    <td width="81%" align="center" bgcolor="#FFFFFF" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#f6f6f6" align="right">ͳ���ߣ�<%=SysAdmin.AdminTitle%>��<%="ͳ��ʱ�䣺<b>"&Now&"</b>"%></td>
        </tr>
      </table>
      <table width="75%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="10"></td>
        </tr>
      </table>
<%
Select Case Work
    Case "Static01"
        Call Static01()
    Case "Static02"
        Call Static02()
    Case "Static03"
        Call Static03()
    Case "Static04"
        Call Static04()
    Case "Static05"
        Call Static05()
    Case Else
		Call Static00()
End Select
        
%>
      <%
'//��ʾ
Sub Static00()
%>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" height="35">
        <tr> 
          <td valign="middle" width="10%" height="88"> </td>
          <td valign="middle" width="90%" height="88"> 
            <li><font color="#0000FF">��ѡ������ͳ������</font></li>
            <li>ͳ��ʱ�佫�������ݿ����������,�����ĵ���</li>
            <li>�����ֽű���ʱ������ʾ,���ӳ�IIS�ų�ʱʱ��ֵ</li>
            <li>����ͳ�Ʋ����뾡�����������ݿ���ʸ߷����ڼ����</li>
          </td>
        </tr>
      </table>
      <%
End Sub
%>
      <%
'//��Ҫͳ��
Sub Static01()
    Dim Sql
        Sql="Select Count(*) From News"
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    Dim TotalNewsNum
        TotalNewsNum=Rs(0)
    Rs.Close
    Sql="Select Count(*) From News Where IsChecked=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim UnCheckedNum
        UnCheckedNum=Rs(0)
    Rs.Close
    Sql="Select Count(*) From News Where Del=1"
    Set Rs=Conn.ExeCute(Sql)
    Dim DeletedNum
        DeletedNum=Rs(0)
    Rs.Close
    Sql="Select Count(*) From News Where Created=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim UnCreatedNum
        UnCreatedNum=Rs(0)
    Rs.Close
    '�������
    Sql="Select Count(*) From News Where DateDiff('m',#"&Now()&"#,AddTime)=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim MonthNum
        MonthNum=Rs(0)
    Rs.Close
    '�������
    Sql="Select Count(*) From News Where DateDiff('d',#"&Now()&"#,AddTime)=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim TodayNum
        TodayNum=Rs(0)
    Rs.Close
    '�����޸�
    Sql="Select Count(*) From News Where DateDiff('m',#"&Now()&"#,UpTime)=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim Mdy_MonthNum
        Mdy_MonthNum=Rs(0)
    Rs.Close
    '�����޸�
    Sql="Select Count(*) From News Where DateDiff('d',#"&Now()&"#,UpTime)=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim Mdy_TodayNum
        Mdy_TodayNum=Rs(0)
    Rs.Close
%>
      <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
        <tr> 
          <td width="24%" bgcolor="#FFFFFF" class="BarTitle"><font color="#FF0000">��Դ����[<%=FormatNumber(TotalNewsNum,0,-1)%>]��</font></td>
          <td width="76%" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FF9900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">δ��˵�[<%=FormatNumber(UnCheckedNum,0,-1)%>]��</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (UnCheckedNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#666699">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">����վ��[<%=FormatNumber(DeletedNum,0,-1)%>]��</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (DeletedNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#666633">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">δ �� ��[<%=FormatNumber(UnCheckedNum,0,-1)%>]��</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (UnCreatedNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#006699">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">�������[<%=FormatNumber(MonthNum,0,-1)%>]��</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (MonthNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#009900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">�������[<%=FormatNumber(TodayNum,0,-1)%>]��</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (TodayNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#0000FF">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">�����޸�[<%=FormatNumber(UnCreatedNum,0,-1)%>]��</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (Mdy_MonthNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#009900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">�����޸�[<%=FormatNumber(Mdy_TodayNum,0,-1)%>]��</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (Mdy_TodayNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#0000FF">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <%
End Sub
%>
      <%
'//��Աͳ��
Sub Static02()
    Dim TotalNum
    Dim Sql,Rs
    '������
    Set Rs=Conn.ExeCute("Select Count(*) From Admin")
    TotalNum=Rs(0)
    Rs.Close
    'ͳ�Ƹ���ɫ������
    Set Rs=Conn.ExeCute("Select B.Title,(Select Count(*) From Admin A Where A.Role=B.Id) As AdminNum From Admin_Role B")
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2" class="ContentTabBg">
        <tr> 
          <td height="2" width="24%" class="BarTitle"><font color="#FF0000">��Ա����[<%=TotalNum%>]��</font></td>
          <td height="2" width="76%" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FF9900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%
    While Not Rs.Eof
%>
        <tr> 
          <td height="2" width="24%" class="BarTitle"><%=Rs("Title")%>[<%=Rs("AdminNum")%>]��</td>
          <td height="2" width="76%" bgcolor="#FFFFFF"> 
            <table width="<%If TotalNum=0 Then Response.Write "0%" Else Response.Write (Rs("AdminNum")/TotalNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#666699">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%
        Rs.MoveNext
    Wend
%>
      </table>
      <%
    Rs.Close
End Sub
%>
      <%
'//��Դ�ֲ����·�
Sub Static03()
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2" valign="middle"> 
            <form name="form1" method="post" action="?">
              <li> ������Ԥ��������ݣ� 
                <input type="hidden" name="Work" value="<%=Work%>">
                <input type="text" name="yyyy" value="<%If Request("yyyy")="" Then Response.Write Year(Now()) Else Response.Write Request("yyyy") End If%>" class="Input" size="8" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">
                <input name="Submit35" type="submit" class="button01-out" value="ȷ ��"">
              </li>
            </form>
          </td>
        </tr>
      </table>
      <%
    If Request("yyyy")<>"" And IsNumeric(Request("yyyy")) Then
        Dim Rs,Sql
        Dim TotalNum
        Set Rs=Conn.ExeCute("Select Count(*) From News Where Year(AddTime)="&CInt(Request("yyyy")))
        TotalNum=Rs(0)
        Rs.Close
        Set Rs=Conn.ExeCute("Select Count(*) As Num,Month(AddTime) As MonthTitle From News Where Year(AddTime)="&Cint(Request("yyyy"))&" Group By Month(AddTime)")
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2" class="ContentTabBg">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2"> 
            <li>����Ϊ<b><%=Request("yyyy")%></b>���ڸ��·���Դ��ӵķֲ����</li>
          </td>
        </tr>
        <tr> 
          <td height="1" width="24%" class="BarTitle"><font color="#FF0000"><%=Request("yyyy")&"����["&TotalNum&"]"%>��</font></td>
          <td height="1" width="76%" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FF9900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%
        While Not Rs.Eof
%>
        <tr> 
          <td height="1" width="24%" class="BarTitle"><%=Rs("MonthTitle")%>��[<%=Rs("Num")%>]��</td>
          <td height="1" width="76%" bgcolor="#FFFFFF"> 
            <table width="<%If TotalNum=0 Then Response.Write "0%" Else Response.Write (Rs("Num")/TotalNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#666699">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%
            Rs.MoveNext
        Wend
%>
      </table>
      <%
    End If
End Sub
%>
      <%
'//��Դ�ֲ���Сʱ
Sub Static04()
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2" valign="middle"> 
            <form name="form1" method="post" action="?">
              <li> ������Ԥ��������ݣ� 
                <input type="hidden" name="Work" value="<%=Work%>">
                <input type="text" name="yyyy" value="<%If Request("yyyy")="" Then Response.Write Year(Now()) Else Response.Write Request("yyyy") End If%>" class="Input" size="8" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">
                <select name="mm" class="input">
                  <option value="">ȫ��</option>
<%
                    Dim I
                    For I=1 To 12
                        If Request("mm")=CStr(I) Then
                            Response.Write "<option value="""&I&""" selected>"&I&"��</option>" & vbCrLf
                        Else
                            Response.Write "<option value="""&I&""">"&I&"��</option>" & vbCrLf
                        End If
                    Next
%>
                </select>
                <input name="Submit352" type="submit" class="button01-out" value="ȷ ��"">
              </li>
            </form>
          </td>
        </tr>
      </table>
      <%
    If Request("yyyy")<>"" And IsNumeric(Request("yyyy")) Then
        Dim Rs,Sql
        Dim TotalNum
        Sql="Select Count(*) From News Where Year(AddTime)="&CInt(Request("yyyy"))
        If Request("mm")<>"" Then
            Sql=Sql & " And Month(AddTime)=" & CInt(Request("mm"))
        End If
        Set Rs=Conn.ExeCute(Sql)
        TotalNum=Rs(0)
        Rs.Close
        Sql="Select Count(*) As Num,Hour(AddTime) As HourTitle From News Where Year(AddTime)="&CInt(Request("yyyy"))
        If Request("mm")<>"" Then
            Sql=Sql & " And Month(AddTime)=" & CInt(Request("mm")) & " Group By Hour(AddTime)"
        Else
            Sql=Sql & " Group By Hour(AddTime)"
        End If
        Set Rs=Conn.ExeCute(Sql)
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2" class="ContentTabBg">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2"> 
            <li>����Ϊ<b><%=Request("yyyy")%></b>��<%If Request("mm")<>"" Then Response.Write "<b>"&Request("mm")&"�·�</b>" End If%>�ڸ�Сʱ��Դ��ӵķֲ����</li>
          </td>
        </tr>
        <tr> 
          <td height="1" width="23%" class="BarTitle"><font color="#FF0000"><%=Request("yyyy")&"��"%> 
            <%If Request("mm")<>"" Then Response.Write Request("mm")&"��" End If%>
            <%="��["&TotalNum&"]"%>��</font></td>
          <td height="1" width="77%" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FF9900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
<%
        While Not Rs.Eof
%>
        <tr> 
          <td height="1" width="23%" class="BarTitle"><%=Rs("HourTitle")%>��[<%=Rs("Num")%>]��</td>
          <td height="1" width="77%" bgcolor="#FFFFFF"> 
            <table width="<%If TotalNum=0 Then Response.Write "0%" Else Response.Write (Rs("Num")/TotalNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#009900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%
            Rs.MoveNext
        Wend
%>
      </table>
      <%
    End If
End Sub
%>
      <%
'//��Ŀ�ֲ����
Sub Static05()
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2" valign="middle"> 
            <form name="form1" method="post" action="?">
              <li> ������Ԥ��������ݣ� 
                <input type="hidden" name="Work" value="<%=Work%>">
                <input type="text" name="yyyy" value="<%If Request("yyyy")="" Then Response.Write Year(Now()) Else Response.Write Request("yyyy") End If%>" class="Input" size="8" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">
                <select name="mm" class="input">
                  <option value="">ȫ��</option>
                  <%
                    Dim I
                    For I=1 To 12
                        If Request("mm")=CStr(I) Then
                            Response.Write "<option value="""&I&""" selected>"&I&"��</option>" & vbCrLf
                        Else
                            Response.Write "<option value="""&I&""">"&I&"��</option>" & vbCrLf
                        End If
                    Next
%>
                </select>
                <input name="Submit3522" type="submit" class="button01-out" value="ȷ ��"">
              </li>
            </form>
          </td>
        </tr>
      </table>
      <%
    If Request("yyyy")<>"" And IsNumeric(Request("yyyy")) Then
        Dim Rs,Sql
        Dim TotalNum
        Sql="Select Count(*) From News Where Year(AddTime)="&CInt(Request("yyyy"))
        If Request("mm")<>"" Then
            Sql=Sql & " And Month(AddTime)=" & CInt(Request("mm"))
        End If
        Set Rs=Conn.ExeCute(Sql)
        TotalNum=Rs(0)
        Rs.Close
        Sql="Select Count(*) As Num,(Select Title From ClassList B Where B.Id=A.Class) As ClassTitle From News A Where Year(AddTime)="&CInt(Request("yyyy"))
        If Request("mm")<>"" Then
            Sql=Sql & " And Month(AddTime)=" & CInt(Request("mm")) & " Group By Class"
        Else
            Sql=Sql & " Group By Class"
        End If
        Set Rs=Conn.ExeCute(Sql)
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2" class="ContentTabBg">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2"> 
            <li>����Ϊ<b><%=Request("yyyy")%></b>��<%If Request("mm")<>"" Then Response.Write "<b>"&Request("mm")&"�·�</b>" End If%>��Դ�ڸ���Ŀ�еķֲ����</li>
          </td>
        </tr>
        <tr> 
          <td height="1" width="22%" class="BarTitle"><font color="#FF0000"><%=Request("yyyy")&"��"%> 
            <%If Request("mm")<>"" Then Response.Write Request("mm")&"��" End If%>
            <%="��["&TotalNum&"]"%>��</font></td>
          <td height="1" width="78%" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FF9900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%
        While Not Rs.Eof
%>
        <tr> 
          <td height="1" width="22%" class="BarTitle"><%=Rs("ClassTitle")%>[<%=Rs("Num")%>]��</td>
          <td height="1" width="78%" bgcolor="#FFFFFF"> 
            <table width="<%If TotalNum=0 Then Response.Write "0%" Else Response.Write (Rs("Num")/TotalNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#009900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%
            Rs.MoveNext
        Wend
%>
      </table>
      <%
    End If
End Sub
%>
    </td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="4">
  <tr> 
    <td width="30%" align="center"> 
      <input name="Submit33" type="button" class="button01-out" value="��  ӡ" onClick="window.print()">
      <input name="Submit34" type="button" class="button01-out" value="��  ��" onClick="window.location.reload();">
      <input name="Submit3" type="button" class="button01-out" value="��  ��" onClick="window.history.back();">
    </td>
  </tr>
</table>
</body>
</html>
