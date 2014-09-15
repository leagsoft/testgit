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
'//生成下载中心首页内容
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
        UpdateOk()
End Select

'//更新成功提示
Function UpdateOk()
    Response.Write("<script>alert(""<操作成功>\n页面更新成功"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End Function

'//今日报道
Function Update01()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../down/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 2 Title,ClassTitle2,ClassUrl,FilePath,SmallImg,ShortContent,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',27,')<>0 Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    If Rs.Eof And Rs.Bof Then
        Exit Function
    End If

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><img src=""" & Rs("SmallImg") & """ width=""100"" height=""100"" border=""1"" style=""border-color:#000000""></a>"

    TClass.OpenTemplate(TemplateFilePath)
    With TClass
        .StartElement="<!---TodayPic01:start-->"
        .EndElement="<!---TodayPic01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><strong>" & Rs("Title") & "</strong></a><br>" & StrClass.CutStr(Rs("ShortContent"),150)

    With TClass
        .StartElement="<!---Today01:start-->"
        .EndElement="<!---Today01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    If Rs.Eof Then
        Exit Function
    Else
        Rs.MoveNext
    End If

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><img src=""" & Rs("SmallImg") & """ width=""100"" height=""100"" border=""1"" style=""border-color:#000000""></a>"

    With TClass
        .StartElement="<!---TodayPic02:start-->"
        .EndElement="<!---TodayPic02:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><strong>" & Rs("Title") & "</strong></a><br>" & StrClass.CutStr(Rs("ShortContent"),150)

    With TClass
        .StartElement="<!---Today02:start-->"
        .EndElement="<!---Today02:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    TClass.Save()
    Set TClass=Nothing

    Rs.Close
    Set Rs=Nothing
End Function

'//推荐下载
Function Update02()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../down/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,ClassTitle2,ClassUrl,FilePath,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',18,')<>0  Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & "<a href=""" & Rs("ClassUrl") & """ target=""_blank"">" & Rs("ClassTitle2") & "</a>" &  "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
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

'//客户端下载
Function Update03()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../down/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,ClassTitle2,ClassUrl,FilePath,AddTime From view_NewsInfo Where Class In (53" & AllChildClass(53) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & "<a href=""" & Rs("ClassUrl") & """ target=""_blank"">" & Rs("ClassTitle2") & "</a>" &  "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--DownList01:start-->"
        .EndElement="<!--DownList01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//补丁下载 
Function Update04()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../down/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,ClassTitle2,ClassUrl,FilePath,AddTime From view_NewsInfo Where Class In (54" & AllChildClass(54) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & "<a href=""" & Rs("ClassUrl") & """ target=""_blank"">" & Rs("ClassTitle2") & "</a>" &  "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--DownList02:start-->"
        .EndElement="<!--DownList02:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//游戏试玩  
Function Update05()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../down/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,ClassTitle2,ClassUrl,FilePath,AddTime From view_NewsInfo Where Class In (55" & AllChildClass(55) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & "<a href=""" & Rs("ClassUrl") & """ target=""_blank"">" & Rs("ClassTitle2") & "</a>" &  "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--DownList03:start-->"
        .EndElement="<!--DownList03:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//常用工具
Function Update06()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../down/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,ClassTitle2,ClassUrl,FilePath,AddTime From view_NewsInfo Where Class In (56" & AllChildClass(56) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & "<a href=""" & Rs("ClassUrl") & """ target=""_blank"">" & Rs("ClassTitle2") & "</a>" &  "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--DownList04:start-->"
        .EndElement="<!--DownList04:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//下载排行
Function Update07()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../down/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 8 Title,ClassTitle2,ClassUrl,FilePath,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',28,')<>0  Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"·[" & "<a href=""" & Rs("ClassUrl") & """ target=""_blank"">" & Rs("ClassTitle2") & "</a>" &  "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--DownTopList:start-->"
        .EndElement="<!--DownTopList:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function
%>