<!--#include file="../Include/Conn.asp" -->
<!--#include file="../Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="../Include/Config.asp" -->
<!--#include file="../Include/ClassList_Fun.asp" -->
<!--#include file="../Include/CreateFile_Fun.asp" -->
<!--#include file="../Include/Tkl_StringClass.asp" -->
<!--#include file="../Include/Tkl_TemplateClass.asp" -->
<!--#Include File="../Include/OnlineClass.asp" -->
<!--#Include File="../Include/UpdateAdminTime.asp" -->
<%
'////////////////////////////////////////////////////////////////////
'//本页：
'//生成default.htm首页内容
'////////////////////////////////////////////////////////////////////

Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If

If Not SysAdmin.UpdatePage Then
    Response.Write("<script>alert(""<操作失败>\n你的权限不足"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End If

Call UpdateAdminTime()

Dim StrClass
Set StrClass = New Tkl_StringClass

Select Case Request("Work")
    Case "Update01" :
        Update01()
        UpdateOk()
    Case "All" :
        Update01()
        Update02()
        Update03()
        Update04()
        Update05()
        Update06()
        Update07()
        Update08()
        UpdateOk()
End Select

'//更新成功提示
Function UpdateOk()
    Response.Write("<script>alert(""<操作成功>\n页面更新成功"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End Function

'//新闻动态
Function Update01()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,FilePath,AddTime From view_NewsInfo Where Class In (34" & AllChildClass(34) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--TopNews:start-->"
        .EndElement="<!--TopNews:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//在线游戏
Function Update02()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,FilePath,AddTime From view_NewsInfo Where Class In (45" & AllChildClass(45) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--OnlineGame:start-->"
        .EndElement="<!--OnlineGame:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//攻略秘技
Function Update03()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,FilePath,AddTime From view_NewsInfo Where Class In (42" & AllChildClass(42) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--Glmj:start-->"
        .EndElement="<!--Glmj:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//玩家风彩
Function Update04()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,FilePath,AddTime From view_NewsInfo Where Class In (47" & AllChildClass(47) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--Gamer:start-->"
        .EndElement="<!--Gamer:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//游戏评论
Function Update05()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,FilePath,AddTime From view_NewsInfo Where Class In (49" & AllChildClass(49) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--GameComment:start-->"
        .EndElement="<!--GameComment:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//图片中心
Sub Update06()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 4 Title,FilePath,SmallImg,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',16,')<>0 And Class In (39" & AllChildClass(39) & ") Order By Id DESC"
	Set Rs=Conn.ExeCute(Sql)
	While Not Rs.Eof
		strHtml=strHtml&"<td><a href="""&Rs("FilePath")&""" target=""_blank""><img src="""&Rs("SmallImg")&""" width=""120"" height=""90"" border=""1"" style=""border-color:#000000""></a></td>" & vbCrLf
		Rs.MoveNext
	Wend
	strHtml="<table width=""100%""  border=""0"" cellspacing=""0"" height=""120"" cellpadding=""2""><tr align=""center"" valign=""middle"">" & strHtml & "</tr></table>"
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--PicCenter:start-->"
        .EndElement="<!--PicCenter:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Sub

'//下载中心 － 更新
Function Update07()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,FilePath,AddTime From view_NewsInfo Where Class In (52" & AllChildClass(52) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--DownNew:start-->"
        .EndElement="<!--DownNew:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//下载中心 － 推荐
Function Update08()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,FilePath,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',18,')<>0  Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--DownRecommand:start-->"
        .EndElement="<!--DownRecommand:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function
%>