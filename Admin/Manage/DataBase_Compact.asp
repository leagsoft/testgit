<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If

%>
<html>
<head>
<title>DataBase_Compact</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link href="Include/ManageStyle.css" rel="stylesheet" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>

<body>
<%
Dim Work
    Work=Request("Work")
Select Case Work
    Case "DoCompact"
        Call DoCompact()
    Case "BakDB"
        Call BakDB()
    Case "DoBakDB"
        Call DoBakDB()
    Case "ReBakDB"
        Call ReBakDB()
    Case "ExeCuteSql"
        Call ExeCuteSql()
    Case "DoExeCuteSql"
        Call DoExeCuteSql()
    Case Else
        Call CompactDB()
End Select
%>
<%Sub CompactDB()%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td align="center" class="BarTitleBg">压缩 数据库</td>
  </tr>
  <tr>
    <td height="105" align="center" class="BarTitle"><form name="form1" method="post" action="?">
        <table width="75%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="32" align="center" class="BarText">注：定期对Access数据进行压缩操作,可以极大地减小库文件大小,提高数据读取性能...<br>
              建意在压缩数据库前对数据库进行一次<a href="?Work=BakDB">备份</a>！</td>
          </tr>
          <tr>
            <td align="center"><input name="Submit" type="submit" class="button01-out" value="确  定">
              <input name="Submit3" type="button" class="button01-out" value="返  回" onclick="window.history.back();">
              <input name="Work" type="hidden" id="Work4" value="DoCompact"> </td>
          </tr>
        </table>
        </form></td>
  </tr>
</table>
<%End Sub%>
<%Sub BakDB()%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td align="center" bgcolor="#CCCCCC" class="BarTitleBg">备份 数据库</td>
  </tr>
  <tr> 
    <td height="105" align="center" bgcolor="#F6f6f6"><form action="?" method="post" name="form2" id="form2">
        <table width="75%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="27" align="center" class="BarText">注：系统自动将资源数据库备份为原数据库同目录中名为 <font color="#006600">DataBase.mdb.bak</font> 
              的文件,同时会复盖上次的备份文件</td>
          </tr>
          <tr>
            <td align="center"><input name="Submit2" type="submit" class="button01-out" value="确  定"> 
              <input name="Submit32" type="button" class="button01-out" value="返  回" onclick="window.history.back();"> 
              <input name="Work" type="hidden" id="Work22" value="DoBakDB"> 
            </td>
          </tr>
        </table>
        
      </form></td>
  </tr>
</table>
<%End Sub%>
<%Sub ReBakDB()%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td align="center" bgcolor="#CCCCCC" class="BarTitleBg">还原 数据库</td>
  </tr>
  <tr> 
    <td height="105" align="center" bgcolor="#F6f6f6"><table width="75%" height="80" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="35" align="center" class="BarText">为确保数据完整性,系统暂不提供此操作！请你手动进行还原<br>
            方法：将数据库同级目录中的.bak数据加文件改名为.mdb既可</td>
        </tr>
        <tr>
          <td align="center"> 
            <input name="Submit33" type="button" class="button01-out" value="返  回" onclick="window.history.back();"></td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
<%End Sub%>
<%Sub ExeCuteSql()%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td align="center" bgcolor="#CCCCCC" class="BarTitleBg">执行Sql脚本</td>
  </tr>
  <tr> 
    <td height="105" align="center" bgcolor="#F6f6f6" class="BarTitle"><form action="?" method="post" name="form4" id="form4">
        <table width="90%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="16" class="BarText">注：一次只能执行一条Sql语句</td>
          </tr>
          <tr>
            <td height="300" align="center" valign="top"> 
                <textarea name="Content" rows="5" wrap="OFF" class="Input" id="Content" style="width:100%;height:100%"></textarea>          
              </td>
          </tr>
          <tr> 
            <td align="center"><input name="Submit4" type="submit" class="button01-out" value="确  定"> 
              <input name="Submit22" type="reset" class="button01-out" value="还  原"> 
              <input name="Submit34" type="button" class="button01-out" value="返  回" onclick="window.history.back();"> 
              <input name="Work" type="hidden" id="Work" value="DoExeCuteSql"> 
            </td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<%End Sub%>
</body>
</html>
<%
'////////////////////////////
'//压缩数据库
Sub DoCompact()
    If Not SysAdmin.ManageDataBase Then
        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim I
    Dim TargetDB,ResourceDB
    Dim oJetEngine
    Dim Fso
    Const Jet_Conn_Partial = "Provider=Microsoft.Jet.OLEDB.4.0; Data source="
    Set oJetEngine = Server.CreateObject("JRO.JetEngine")
    Set Fso= Server.CreateObject(FsoObjectStr)
    '关闭数据库链接
    Conn.Close
    Set Conn=Nothing

    For I=1 To UBound(DBName)
        ResourceDB=Server.MapPath(DBName(I))

        If Fso.FileExists(ResourceDB) Then
            '建立临时文件
            TargetDB=Server.MapPath(DBName(I)&".bak")
            If Fso.FileExists(TargetDB) Then
                Fso.DeleteFile(TargetDB)
            End If
            oJetEngine.CompactDatabase Jet_Conn_Partial&ResourceDB,Jet_Conn_Partial&TargetDB
            Fso.DeleteFile ResourceDB
            Fso.MoveFile TargetDB,ResourceDB
        End If
    Next

    Set Fso=Nothing
    Set oJetEngine=Nothing

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "压缩数据库")
    Set LogClass=Nothing

    Response.Write("<script>alert(""<操作成功>\n数据库压缩完成"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End Sub

'////////////////////////////
'//备份数据库
Sub DoBakDB()
    If Not SysAdmin.ManageDataBase Then
        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim I
    Dim TargetDB,ResourceDB
    Dim Fso
    Set Fso= Server.CreateObject(FsoObjectStr)
    '关闭数据库链接
    Conn.Close
    Set Conn=Nothing

    For I=1 To UBound(DBName)
        ResourceDB=Server.MapPath(DBName(I))
        If Fso.FileExists(ResourceDB) Then
            '原数据库文是否存在
            If Fso.FileExists(ResourceDB) Then
                '建立原备份文件
                TargetDB=Server.MapPath(DBName(I)&".bak")
                Fso.CopyFile ResourceDB,TargetDB,True
            End If
        End If
    Next

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "备份数据库")
    Set LogClass=Nothing

    Set Fso=Nothing
    Response.Write("<script>alert(""<操作成功>\n数据库备份完成"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End Sub

'////////////////////////////
'//执行Sql操作
Sub DoExeCuteSql()
    If Not SysAdmin.ManageDataBase Then
        Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Content
        Content=Request("Content")
    Err.Clear
    On Error Resume Next
    Conn.BeginTrans
    Conn.ExeCute(Content)
    If Err.Number<>0 Then
        Conn.Rollback
        Response.Write("<script>alert(""<操作失败>\nSql语句有误，请返回查看"& SoftCopyright_Script &""");window.history.back();</script>")

    Else
        Conn.CommitTrans
        Response.Write("<script>alert(""<操作成功>\nSql语句执行完毕"& SoftCopyright_Script &""");window.history.back();</script>")
    End If

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "执行Sql语句")
    Set LogClass=Nothing

    Response.End
End Sub
%>