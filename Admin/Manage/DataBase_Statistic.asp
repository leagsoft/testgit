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
        STitle="简要统计"
    Case "Static02"
        STitle="人员统计"
    Case "Static03"
        STitle="资源分布·年份"
    Case "Static04"
        STitle="资源分布·小时"
    Case "Static05"
        STitle="资源分布·栏目"
    Case Else
        STitle="简要统计"
End Select
%>
<table width="100%" height="192" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr align="center"> 
    <td height="10" colspan="2" class="BarTitleBg">“<%=STitle%>”数据库统计 </td>
  </tr>
  <tr align="center" bgcolor="#FFFFCC"> 
    <td height="10" colspan="2"> 
      <input name="Submit32" type="button" class="button01-out" value="简要统计" onClick="SelStatic(1)">
      <input name="Submit322" type="button" class="button01-out" value="人员统计" onClick="SelStatic(2)" title="系统将分析所有系统帐户在各角色中的分布情况">
      <input name="Submit323" type="button" class="button02-out" value="月份添加分布" onClick="SelStatic(3)" title="系统将分析指定年度中各月份资源的添加分布情况">
      <input name="Submit3232" type="button" class="button02-out" value="小时添加分布" onClick="SelStatic(4)" title="系统将分析指定年度中各小时资源的添加分布情况">
      <input name="Submit32322" type="button" class="button02-out" value="栏目资源分布" onClick="SelStatic(5)" title="系统将分析指定年度中栏目资源的添加分布情况">
    </td>
  </tr>
  <tr> 
    <td width="19%" height="168" align="center" valign="top" bgcolor="#FFFFFF"><img src="Images/Manage/Statistic.gif" width="136" height="125"></td>
    <td width="81%" align="center" bgcolor="#FFFFFF" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#f6f6f6" align="right">统计者：<%=SysAdmin.AdminTitle%>　<%="统计时间：<b>"&Now&"</b>"%></td>
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
'//提示
Sub Static00()
%>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" height="35">
        <tr> 
          <td valign="middle" width="10%" height="88"> </td>
          <td valign="middle" width="90%" height="88"> 
            <li><font color="#0000FF">请选择数据统计类型</font></li>
            <li>统计时间将随着数据库增大而变慢,请耐心等侍</li>
            <li>若出现脚本超时错误提示,请延长IIS脚超时时间值</li>
            <li>数据统计操作请尽量避免在数据库访问高峰期期间进行</li>
          </td>
        </tr>
      </table>
      <%
End Sub
%>
      <%
'//简要统计
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
    '本月添加
    Sql="Select Count(*) From News Where DateDiff('m',#"&Now()&"#,AddTime)=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim MonthNum
        MonthNum=Rs(0)
    Rs.Close
    '本日添加
    Sql="Select Count(*) From News Where DateDiff('d',#"&Now()&"#,AddTime)=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim TodayNum
        TodayNum=Rs(0)
    Rs.Close
    '本月修改
    Sql="Select Count(*) From News Where DateDiff('m',#"&Now()&"#,UpTime)=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim Mdy_MonthNum
        Mdy_MonthNum=Rs(0)
    Rs.Close
    '本日修改
    Sql="Select Count(*) From News Where DateDiff('d',#"&Now()&"#,UpTime)=0"
    Set Rs=Conn.ExeCute(Sql)
    Dim Mdy_TodayNum
        Mdy_TodayNum=Rs(0)
    Rs.Close
%>
      <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
        <tr> 
          <td width="24%" bgcolor="#FFFFFF" class="BarTitle"><font color="#FF0000">资源总数[<%=FormatNumber(TotalNewsNum,0,-1)%>]：</font></td>
          <td width="76%" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FF9900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">未审核的[<%=FormatNumber(UnCheckedNum,0,-1)%>]：</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (UnCheckedNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#666699">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">回收站内[<%=FormatNumber(DeletedNum,0,-1)%>]：</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (DeletedNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#666633">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">未 生 成[<%=FormatNumber(UnCheckedNum,0,-1)%>]：</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (UnCreatedNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#006699">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">本月添加[<%=FormatNumber(MonthNum,0,-1)%>]：</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (MonthNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#009900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">本日添加[<%=FormatNumber(TodayNum,0,-1)%>]：</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (TodayNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#0000FF">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">本月修改[<%=FormatNumber(UnCreatedNum,0,-1)%>]：</td>
          <td bgcolor="#FFFFFF" width="76%"> 
            <table width="<%If TotalNewsNum=0 Then Response.Write "0%" Else Response.Write (Mdy_MonthNum/TotalNewsNum*100)&"%" End If%>" border="0" cellpadding="0" cellspacing="0" bgcolor="#009900">
              <tr> 
                <td align="center" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" class="BarTitle" width="24%">本日修改[<%=FormatNumber(Mdy_TodayNum,0,-1)%>]：</td>
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
'//人员统计
Sub Static02()
    Dim TotalNum
    Dim Sql,Rs
    '总人数
    Set Rs=Conn.ExeCute("Select Count(*) From Admin")
    TotalNum=Rs(0)
    Rs.Close
    '统计各角色的人数
    Set Rs=Conn.ExeCute("Select B.Title,(Select Count(*) From Admin A Where A.Role=B.Id) As AdminNum From Admin_Role B")
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2" class="ContentTabBg">
        <tr> 
          <td height="2" width="24%" class="BarTitle"><font color="#FF0000">人员总数[<%=TotalNum%>]：</font></td>
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
          <td height="2" width="24%" class="BarTitle"><%=Rs("Title")%>[<%=Rs("AdminNum")%>]：</td>
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
'//资源分布·月份
Sub Static03()
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2" valign="middle"> 
            <form name="form1" method="post" action="?">
              <li> 请输入预分析的年份： 
                <input type="hidden" name="Work" value="<%=Work%>">
                <input type="text" name="yyyy" value="<%If Request("yyyy")="" Then Response.Write Year(Now()) Else Response.Write Request("yyyy") End If%>" class="Input" size="8" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">
                <input name="Submit35" type="submit" class="button01-out" value="确 定"">
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
            <li>以下为<b><%=Request("yyyy")%></b>年在各月份资源添加的分布情况</li>
          </td>
        </tr>
        <tr> 
          <td height="1" width="24%" class="BarTitle"><font color="#FF0000"><%=Request("yyyy")&"年总["&TotalNum&"]"%>：</font></td>
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
          <td height="1" width="24%" class="BarTitle"><%=Rs("MonthTitle")%>月[<%=Rs("Num")%>]：</td>
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
'//资源分布·小时
Sub Static04()
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2" valign="middle"> 
            <form name="form1" method="post" action="?">
              <li> 请输入预分析的年份： 
                <input type="hidden" name="Work" value="<%=Work%>">
                <input type="text" name="yyyy" value="<%If Request("yyyy")="" Then Response.Write Year(Now()) Else Response.Write Request("yyyy") End If%>" class="Input" size="8" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">
                <select name="mm" class="input">
                  <option value="">全年</option>
<%
                    Dim I
                    For I=1 To 12
                        If Request("mm")=CStr(I) Then
                            Response.Write "<option value="""&I&""" selected>"&I&"月</option>" & vbCrLf
                        Else
                            Response.Write "<option value="""&I&""">"&I&"月</option>" & vbCrLf
                        End If
                    Next
%>
                </select>
                <input name="Submit352" type="submit" class="button01-out" value="确 定"">
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
            <li>以下为<b><%=Request("yyyy")%></b>年<%If Request("mm")<>"" Then Response.Write "<b>"&Request("mm")&"月份</b>" End If%>在各小时资源添加的分布情况</li>
          </td>
        </tr>
        <tr> 
          <td height="1" width="23%" class="BarTitle"><font color="#FF0000"><%=Request("yyyy")&"年"%> 
            <%If Request("mm")<>"" Then Response.Write Request("mm")&"月" End If%>
            <%="总["&TotalNum&"]"%>：</font></td>
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
          <td height="1" width="23%" class="BarTitle"><%=Rs("HourTitle")%>点[<%=Rs("Num")%>]：</td>
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
'//栏目分布情况
Sub Static05()
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr bgcolor="#FFFFFF"> 
          <td height="0" colspan="2" valign="middle"> 
            <form name="form1" method="post" action="?">
              <li> 请输入预分析的年份： 
                <input type="hidden" name="Work" value="<%=Work%>">
                <input type="text" name="yyyy" value="<%If Request("yyyy")="" Then Response.Write Year(Now()) Else Response.Write Request("yyyy") End If%>" class="Input" size="8" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">
                <select name="mm" class="input">
                  <option value="">全年</option>
                  <%
                    Dim I
                    For I=1 To 12
                        If Request("mm")=CStr(I) Then
                            Response.Write "<option value="""&I&""" selected>"&I&"月</option>" & vbCrLf
                        Else
                            Response.Write "<option value="""&I&""">"&I&"月</option>" & vbCrLf
                        End If
                    Next
%>
                </select>
                <input name="Submit3522" type="submit" class="button01-out" value="确 定"">
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
            <li>以下为<b><%=Request("yyyy")%></b>年<%If Request("mm")<>"" Then Response.Write "<b>"&Request("mm")&"月份</b>" End If%>资源在各栏目中的分布情况</li>
          </td>
        </tr>
        <tr> 
          <td height="1" width="22%" class="BarTitle"><font color="#FF0000"><%=Request("yyyy")&"年"%> 
            <%If Request("mm")<>"" Then Response.Write Request("mm")&"月" End If%>
            <%="总["&TotalNum&"]"%>：</font></td>
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
          <td height="1" width="22%" class="BarTitle"><%=Rs("ClassTitle")%>[<%=Rs("Num")%>]：</td>
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
      <input name="Submit33" type="button" class="button01-out" value="打  印" onClick="window.print()">
      <input name="Submit34" type="button" class="button01-out" value="更  新" onClick="window.location.reload();">
      <input name="Submit3" type="button" class="button01-out" value="返  回" onClick="window.history.back();">
    </td>
  </tr>
</table>
</body>
</html>
