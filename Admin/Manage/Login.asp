<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/CfsEnCode.asp" -->
<!--#include file="CheckAdmin.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#Include File="Include/OnlineClass.asp" -->
<!--#Include File="Include/UpdateAdminTime.asp" -->
<!--#Include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<html>
<head>
<title>Login.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="CBRCGD,GuangDongGuangZhou,Star_Info">
<meta name="Version" content="Tsys V1.1">
<link rel="stylesheet" href="Include/ManageStyle.css" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
<script src="Include/Tkl_Tooltip.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" onLoad="try{form1.Title.focus()}catch(exception){}">
<%
Dim LogClass
Set LogClass=New Tkl_LogClass
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
Dim Coll
Set Coll = New UserInfo_Collection_Class
If Request("Work")="LogOut" Then
    SysAdmin.LogOut()
    Coll.Remove(SysAdmin.AdminTitle)
    If CBool(Request("CloseWin")) Then
        Response.Write "<script>top.close();</script>"
        Response.End
    End If
    Response.Redirect "?"
End If
If CBool(SysAdmin.Logined) Then
    Call UpdateAdminTime()
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
    Dim Sql
        '**Sql="Select * From View_AdminInfo Where UCase(Title)='" & UCase(SysAdmin.AdminTitle) & "'"
        Sql="Select * From View_AdminInfo Where Title='" & UCase(SysAdmin.AdminTitle) & "'"
        Rs.Open Sql,Conn,1,3
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write ("无法获得管理员信息")
        Response.End
    End If 
    If Request("UpAdminLoginDate_ThisTime")="True" Then
        Rs("LastLoginTime")=Rs("LoginTime")
        Rs("LoginTime")=Now()
        Rs("LoginCount")=Rs("LoginCount")+1
    End If
    Rs.Update
    Dim RedirectPage
        RedirectPage=Request("RedirectPage")
    If RedirectPage<>"" Then
        Response.Cookies("TsysLoginCookie")("RedirectPage")=RedirectPage
        If RedirectPage<>"?" Then
            Rs.Close
            Response.Redirect RedirectPage
        End If
    End If
%>
<table width="85%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr align="center"> 
    <td height="15" colspan="2" class="BarTitleBg">帐户信息</td>
  </tr>
  <tr> 
    <td width="23%" height="284" align="center" valign="middle"><img src="Images/Skin/CBRCGD.gif" width="100" height="252"></td>
    <td width="77%">
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="ContentTabBg">
        <tr bgcolor="#FFFFFF"> 
          <td width="32%" class="BarTitle">帐户:</td>
          <td width="68%" class="BarText"><font size="" color="#0000FF"><%=Rs("Title")%></font></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="32%" class="BarTitle">帐户昵称:</td>
          <td width="68%" class="BarText"><%=Rs("NickName")%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="32%" class="BarTitle">账户角色:</td>
          <td width="68%" class="BarText"><span title="<%=Rs("Content")%>"><%=Rs("RoleTitle")%></span></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="32%" class="BarTitle">创建时间:</td>
          <td width="68%" class="BarText"><%=Rs("AddTime")%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="32%" class="BarTitle">最后登录:</td>
          <td width="68%" class="BarText"><%=Rs("LastLoginTime")%></td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td width="32%" class="BarTitle">登录次数：</td>
          <td width="68%" class="BarText"><%=Rs("LoginCount")%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="32%" class="BarTitle">&nbsp; </td>
          <td width="68%"> 
            <input name="Submit2" type="button" class="button01-out" value="注  消" onClick="window.location='?Work=LogOut'">
            <input name="Submit3" type="button" class="button01-out" value="返  回" onclick="window.history.back();">
          </td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td colspan="2">注：<br>
&nbsp; 1.退出本系统前请先注销用户,否则在系统默认的帐号超时时间内，同名的管理员帐号无法从另一IP登录,同时也防止管理员权限贮留，产生安全瘾患！<br>
&nbsp; 2.本系统要求服务器端及客户端均安装有<strong>IE5.5</strong>以上版本，否则将无法正常使用核心功能。</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<table width="85%" border="0" align="center" cellpadding="2" cellspacing="1" class="ContentTabBg">
  <tr bgcolor="#FFFFFF"> 
    <td colspan="2" bgcolor="#F6f6f6" class="BarTitleBg">·当前在线管理员[<a href="javascript:window.location.reload()"><font color="#FFFF00">刷新</font></a>]</td>
  </tr>
  <tr> 
    <td width="9%" height="42" align="center" valign="top" bgcolor="#FFFFFF" class="BarText"> <img src="Images/Manage/People.gif" width="38" height="38"> 
    </td>
    <td width="91%" bgcolor="#FFFFFF" class="BarText"> &nbsp; 
<%
    Set Coll = New UserInfo_Collection_Class
    Dim TempUserInfo
    Dim I
    For I = 1 To Coll.Count()
        Set TempUserInfo = Coll.GetUser(I)
        Response.Write "<a href=""#"" onmouseover=""showToolTip('ＩＰ地址：" & TempUserInfo.Ip & "<br>登入时间：" & TempUserInfo.AddTime & "<br>最近刷新：" & TempUserInfo.UpTime & "',event.srcElement)"" onmouseout=""hiddenToolTip()"">" & TempUserInfo.Name & "</a>" & "&nbsp;"
    Next
%>
    </td>
  </tr>
</table>
<br>
<%
    Rs.Close
    Set Rs=Nothing
Else
    Dim Title,Pwd
    If Request("Work")="Check" Then        '检验管理员
       If Not ChkEnableLogin() Then 
            Response.Write("<script>alert(""<登录失败>\n你的IP由于登录错误次数过多，已经被禁止。一段时间内将无法再次登录，详情请与[超级管理员]联系"& SoftCopyright_Script &""");window.history.back()</script>")
            Response.End
        End If
		'Response.Write sql
		'Response.End
        Title=Replace(Trim(Request("Title")),"'","''")
        'Response.Write Title
        'Response.End 
        Pwd=Replace(Trim(Request("Pwd")),"'","''")
        If Title<>"" And Pwd<>"" Then

            Dim Result
                Result=CheckAdmin(Title,CfsEnCode(Pwd))
            If Result<>"" Then
                If Result="{LOCK}" Then
                    Response.Write("<script>alert(""<登录失败>\n该用户已被锁定,详情请与[超级管理员]联系"& SoftCopyright_Script &""");</script>")
                Else
                    Dim AdmInfo
                        AdmInfo=Split(Result,vbTab,-1,1)
                    '生成管理员信息
                    SysAdmin.AdminLogined="TRUE"
                    SysAdmin.AdminTitle=AdmInfo(0)
                    SysAdmin.AdminPopedom=AdmInfo(1)
                    SysAdmin.AdminRoleTitle=AdmInfo(2)
                    SysAdmin.AdminNickName=AdmInfo(3)
                    SysAdmin.AdminClassPopedom=AdmInfo(4)
                    SysAdmin.AdminTopClassId=AdmInfo(5)

                    Dim myInfo
                    '添加当前管理员入[在线列表]
                    Set myInfo = New UserInfo_Class
                        myInfo.Id = SysAdmin.AdminTitle
                        myInfo.Name = SysAdmin.AdminTitle
                        myInfo.Ip = Request.ServerVariables("REMOTE_ADDR")
                        myInfo.NickName = SysAdmin.AdminNickName
                        myInfo.AddTime = Now
                        myInfo.UpTime = Now
                        myInfo.Remark=""

                    Set Coll = New UserInfo_Collection_Class
                    If Not DubleOnlineUser Then
                        If Coll.Find(myInfo.Name) Then
                            Dim tempmyInfo
                            Set tempmyInfo= Coll.GetUser(myInfo.Name)
                            If Trim(tempmyInfo.Ip)<>Trim(myInfo.Ip) Then
                                Response.Write("<script>alert(""<登录失败>\n该[用户]当前正在线！因些你无法登录,请更换登入名,或要求对方消注登录\n对方登录时间:"&tempmyInfo.AddTime&"\n最后刷新时间:"&tempmyInfo.AddTime&"\n对方IP:"&tempmyInfo.Ip& SoftCopyright_Script &""");window.history.back();</script>")
                                SysAdmin.LogOut()
                                Response.End
                            End If
                        End If
                    End If
                    Coll.Add(myInfo)

                    If CBool(Request("AutoRemberLoginName")) Then
                        Response.Cookies("TsysLoginCookie")("AdminTitle")=myInfo.Name
                        Response.Cookies("TsysLoginCookie").Expires=Date()+AutoRemberLoginName_ExpiresTime
                    Else
                        Response.Cookies("TsysLoginCookie")("AdminTitle")=""
                        Response.Cookies("TsysLoginCookie").Expires=Date()-1
                    End If

                    LogClass.AddLog(myInfo.Name & "登录系统,IP:" & myInfo.Ip)
                    Conn.ExeCute("Delete From LoginLock Where Title='"&Request.ServerVariables("REMOTE_ADDR")&"'")

                    Response.Redirect "?UpAdminLoginDate_ThisTime=True&RedirectPage="&Request("RedirectPage")
                End If
            Else
                '进行安全登记
                Call RemberLoginWrong()
                Response.Write("<script>alert(""<登录失败>\n[用户]或[密码]错误"& SoftCopyright_Script &""");window.history.back();</script>")
            End If
        End If
    End If
%>
<form name="form1" method="post" action="?">
    <br>
  <table width="85%" border="0" align="center" cellpadding="3" cellspacing="1" class="ContentTabBg">
    <tr align="center"> 
      <td colspan="2" class="BarTitleBg">管理员登录</td>
    </tr>
    <tr> 
      <td width="25%" align="center" valign="middle" bgcolor="#FFFFFF"><img src="Images/Skin/CBRCGD.gif" width="100" height="252"></td>
      <td width="75%" bgcolor="#FFFFFF" class="BarTitle">
<table width="100%" height="76" border="0" align="center" cellpadding="3" cellspacing="1" class="ContentTabBg">
          <tr> 
            <td width="27%" height="2" bgcolor="#FFFFFF" class="BarTitle"> 帐户:</td>
            <td width="73%" height="2" bgcolor="#FFFFFF"> <input type="text" name="Title" maxlength="20" class="Input" size="40" tabindex="1" onkeydown="if(event.keyCode==13)event.keyCode=9" value="<%=Request.Cookies("TsysLoginCookie")("AdminTitle")%>"> 
            </td>
          </tr>
          <tr> 
            <td width="27%" height="10" bgcolor="#FFFFFF" class="BarTitle">密码:</td>
            <td width="73%" height="10" bgcolor="#FFFFFF"> <input type="password" name="Pwd" maxlength="20" class="Input" size="40" tabindex="2"> 
            </td>
          </tr>
          <tr> 
            <td width="27%" height="11" bgcolor="#FFFFFF" class="BarTitle">登录页面</td>
            <td height="11" bgcolor="#FFFFFF"><select name="RedirectPage" id="RedirectPage">
                <option value="?" <%If Request.Cookies("TsysLoginCookie")("RedirectPage")="?" Then Response.Write "Selected" End If%>>默认页</option>
                <option value="News_List.asp" <%If Request.Cookies("TsysLoginCookie")("RedirectPage")="News_List.asp" Then Response.Write "Selected" End If%>>资源列表</option>
                <option value="News_Add.asp?Work=AddReco" <%If Request.Cookies("TsysLoginCookie")("RedirectPage")="News_Add.asp?Work=AddReco" Then Response.Write "Selected" End If%>>添加资源</option>
                <!--<option value="DataBase_Statistic.asp" <%If Request.Cookies("TsysLoginCookie")("RedirectPage")="DataBase_Statistic.asp" Then Response.Write "Selected" End If%>>数据统计</option>
                <option value="Comment_List.asp" <%If Request.Cookies("TsysLoginCookie")("RedirectPage")="Comment_List.asp" Then Response.Write "Selected" End If%>>评论管理</option>-->
                <option value="Class_List.asp" <%If Request.Cookies("TsysLoginCookie")("RedirectPage")="Class_List.asp" Then Response.Write "Selected" End If%>>分类列表</option>
                <option value="Admin_List.asp" <%If Request.Cookies("TsysLoginCookie")("RedirectPage")="Admin_List.asp" Then Response.Write "Selected" End If%>>帐户管理</option>
              </select></td>
          </tr>
          <tr> 
            <td width="27%" height="2" bgcolor="#FFFFFF"> <input type="hidden" name="Work" value="Check">
            </td>
            <td width="73%" height="2" bgcolor="#FFFFFF"> 
              <input name="Submit" type="submit" class="button01-out" id="Submit"  value="确  定"> 
              <input name="Submit32" type="button" class="button01-out" value="还  原" onclick="window.history.back();"></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF"></td>
            <td height="22" bgcolor="#FFFFFF" class="BarText">
            <label for="AutoRemberLoginName"><INPUT TYPE="checkbox" id="AutoRemberLoginName" NAME="AutoRemberLoginName" value="1" <% If IsAutoRemberLoginName Then Response.Write "checked" End If%>>记住登录名</label></td>
          </tr>
          <tr> 
            <td height="22" colspan="2" bgcolor="#FFFFFF"><p>注：<br>
                &nbsp; 1.退出本系统前请先注销用户,否则在系统默认的帐号超时时间内，同名的管理员帐号无法从另一IP登录,同时也防止管理员权限贮留，产生安全瘾患！<br>
                &nbsp; 2.本系统要求服务器端及客户端均安装有<strong>IE5.5</strong>以上版本，否则将无法正常使用核心功能。</p>
            </td>
          </tr>
        </table>	  
	  </td>
    </tr>
  </table>
</form>
<%End If%>
</body>
</html>
<%
Set LogClass=Nothing

'//函数：登录安全登记
Sub RemberLoginWrong()
    LogClass.AddLog("IP:" & Request.ServerVariables("REMOTE_ADDR")& "登录失败")
    If Def_UseLoginPolliceMan Then
    '************************************Modify By BennyLiu:20040311 ***************************************
        'Conn.ExeCute("Insert Into LoginWrongLog (Title,AddTime)Values('"&Request.ServerVariables("REMOTE_ADDR")&"',#"&Now()&"#)")
        Conn.ExeCute("Insert Into LoginWrongLog (Title,AddTime)Values('"&Request.ServerVariables("REMOTE_ADDR")&"','"&Now&"')")
    '********************************************* End Modify **********************************************
        '删除不在监视范围内的安全登记
        '**Conn.ExeCute("Delete From LoginWrongLog  Where DateDiff('s',AddTime,Now())>" & Def_StakeoutTimeRange)
        Conn.ExeCute("Delete From LoginWrongLog  Where DateDiff(s,AddTime,getdate())>" & Def_StakeoutTimeRange)
        Dim Rs
        Set Rs=Conn.ExeCute("Select Count(*) From LoginWrongLog Where Title='"&Request.ServerVariables("REMOTE_ADDR")&"'")
        If Rs(0)>Def_EnableLoginWrong_Number Then
            '错误登录次数超标则封IP
            Conn.ExeCute("Insert Into LoginLock (Title,AddTime)Values('"&Request.ServerVariables("REMOTE_ADDR")&"',#"&Now()&"#)")
            '清空当前IP的安全登记
            Conn.ExeCute("Delete From LoginWrongLog Where Title='"&Request.ServerVariables("REMOTE_ADDR")&"'")
        End If
        Rs.Close
        Set Rs=Nothing
    End If
End Sub

'//函数：登录安全检查
'//返回：Bool(True:允许登录)
Function ChkEnableLogin()
    ChkEnableLogin=False
    If Def_UseLoginPolliceMan Then
        '删除超出监视时间范围的错误登录记录
        'SQL SERVER中不能使用Now()函数
        Conn.ExeCute("Delete From LoginLock Where DateDiff(s,AddTime,getdate())>" & Def_LoginWrongLockTimeRange)
        '**Conn.ExeCute("Delete From LoginLock Where DateDiff('s',AddTime,Now())>" & Def_LoginWrongLockTimeRange)
        Dim Rs
        Set Rs=Conn.ExeCute("Select * From LoginLock Where Title='"&Request.ServerVariables("REMOTE_ADDR")&"'")
        ChkEnableLogin=Rs.Eof And Rs.Bof
        Rs.Close
        Set Rs=Nothing
    Else
        ChkEnableLogin=True
    End If
End Function
%>