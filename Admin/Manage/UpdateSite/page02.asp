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
'//生成新闻中心/news/default.htm首页内容
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

'//新游报道
Function Update01()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../news/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 1 Title,FilePath,SmallImg,ShortContent,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',20,')<>0 And Class In (35" & AllChildClass(35) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    If Rs.Eof And Rs.Bof Then
        Exit Function
    End If

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><img src=""" & Rs("SmallImg") & """ width=""121"" height=""121"" border=""1"" style=""border-color:#000000""></a>"

    TClass.OpenTemplate(TemplateFilePath)
    With TClass
        .StartElement="<!---TopNewsImg01:start-->"
        .EndElement="<!---TopNewsImg01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><strong>" & Rs("Title") & "</strong></a><br>" & StrClass.CutStr(Rs("ShortContent"),200)

    With TClass
        .StartElement="<!---TopNews01:start-->"
        .EndElement="<!---TopNews01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With
    TClass.Save()
    Set TClass=Nothing

    Rs.Close
    Set Rs=Nothing
End Function

'//业界动态
Function Update02()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../news/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 1 Title,FilePath,SmallImg,ShortContent,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',20,')<>0 And Class In (36" & AllChildClass(36) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    If Rs.Eof And Rs.Bof Then
        Exit Function
    End If

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><img src=""" & Rs("SmallImg") & """ width=""121"" height=""121"" border=""1"" style=""border-color:#000000""></a>"

    TClass.OpenTemplate(TemplateFilePath)
    With TClass
        .StartElement="<!---TopNewsImg01:start-->"
        .EndElement="<!---TopNewsImg01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><strong>" & Rs("Title") & "</strong></a><br>" & StrClass.CutStr(Rs("ShortContent"),200)

    With TClass
        .StartElement="<!---TopNews01:start-->"
        .EndElement="<!---TopNews01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With
    TClass.Save()
    Set TClass=Nothing

    Rs.Close
    Set Rs=Nothing
End Function

'//杂谈赏析
Function Update03()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../news/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 1 Title,FilePath,SmallImg,ShortContent,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',20,')<>0 And Class In (37" & AllChildClass(37) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    If Rs.Eof And Rs.Bof Then
        Exit Function
    End If

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><img src=""" & Rs("SmallImg") & """ width=""121"" height=""121"" border=""1"" style=""border-color:#000000""></a>"

    TClass.OpenTemplate(TemplateFilePath)
    With TClass
        .StartElement="<!---TopNewsImg03:start-->"
        .EndElement="<!---TopNewsImg03:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><strong>" & Rs("Title") & "</strong></a><br>" & StrClass.CutStr(Rs("ShortContent"),200)

    With TClass
        .StartElement="<!---TopNews03:start-->"
        .EndElement="<!---TopNews03:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With
    TClass.Save()
    Set TClass=Nothing

    Rs.Close
    Set Rs=Nothing
End Function

'//赛事战报
Function Update04()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../news/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 1 Title,FilePath,SmallImg,ShortContent,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',20,')<>0 And Class In (38" & AllChildClass(38) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    If Rs.Eof And Rs.Bof Then
        Exit Function
    End If

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><img src=""" & Rs("SmallImg") & """ width=""121"" height=""121"" border=""1"" style=""border-color:#000000""></a>"

    TClass.OpenTemplate(TemplateFilePath)
    With TClass
        .StartElement="<!---TopNewsImg04:start-->"
        .EndElement="<!---TopNewsImg04:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = "<a href=""" & Rs("FilePath") & """ target=""_blank""><strong>" & Rs("Title") & "</strong></a><br>" & StrClass.CutStr(Rs("ShortContent"),200)

    With TClass
        .StartElement="<!---TopNews04:start-->"
        .EndElement="<!---TopNews04:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With
    TClass.Save()
    Set TClass=Nothing

    Rs.Close
    Set Rs=Nothing
End Function

'//新游报道
Function Update05()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../news/default.htm")
    Dim TClass,strHtml,I
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 16 Title,FilePath,AddTime From view_NewsInfo Where Class In (35" & AllChildClass(35) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)

    TClass.OpenTemplate(TemplateFilePath)

    strHtml = ""
    For I=1 To 8
        If Rs.Eof And Rs.Bof Then
            Exit For
        End If
        strHtml = strHtml & "·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!---TopNewsList01:start-->"
        .EndElement="<!---TopNewsList01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = ""
    For I=1 To 8
        If Rs.Eof And Rs.Bof Then
            Exit For
        End If
        strHtml = strHtml & "·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!---TopNewsList02:start-->"
        .EndElement="<!---TopNewsList02:end-->"
        .Value=strHtml
        .ReplaceTemplate()        
    End With

    Rs.Close
    Set Rs=Nothing

    TClass.Save()
    Set TClass=Nothing
End Function

'//业界动态
Function Update06()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../news/default.htm")
    Dim TClass,strHtml,I
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 16 Title,FilePath,AddTime From view_NewsInfo Where Class In (36" & AllChildClass(36) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)

    TClass.OpenTemplate(TemplateFilePath)

    strHtml = ""
    For I=1 To 8
        If Rs.Eof And Rs.Bof Then
            Exit For
        End If
        strHtml = strHtml & "·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!---TopNewsList03:start-->"
        .EndElement="<!---TopNewsList03:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = ""
    For I=1 To 8
        If Rs.Eof And Rs.Bof Then
            Exit For
        End If
        strHtml = strHtml & "·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!---TopNewsList04:start-->"
        .EndElement="<!---TopNewsList04:end-->"
        .Value=strHtml
        .ReplaceTemplate()        
    End With

    Rs.Close
    Set Rs=Nothing

    TClass.Save()
    Set TClass=Nothing
End Function

'//杂谈赏析
Function Update07()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../news/default.htm")
    Dim TClass,strHtml,I
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 16 Title,FilePath,AddTime From view_NewsInfo Where Class In (37" & AllChildClass(37) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)

    TClass.OpenTemplate(TemplateFilePath)

    strHtml = ""
    For I=1 To 8
        If Rs.Eof And Rs.Bof Then
            Exit For
        End If
        strHtml = strHtml & "·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!---TopNewsList05:start-->"
        .EndElement="<!---TopNewsList05:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = ""
    For I=1 To 8
        If Rs.Eof And Rs.Bof Then
            Exit For
        End If
        strHtml = strHtml & "·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!---TopNewsList06:start-->"
        .EndElement="<!---TopNewsList06:end-->"
        .Value=strHtml
        .ReplaceTemplate()        
    End With

    Rs.Close
    Set Rs=Nothing

    TClass.Save()
    Set TClass=Nothing
End Function

'//赛事战报
Function Update08()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../news/default.htm")
    Dim TClass,strHtml,I
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 16 Title,FilePath,AddTime From view_NewsInfo Where Class In (38" & AllChildClass(38) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)

    TClass.OpenTemplate(TemplateFilePath)

    strHtml = ""
    For I=1 To 8
        If Rs.Eof And Rs.Bof Then
            Exit For
        End If
        strHtml = strHtml & "·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!---TopNewsList07:start-->"
        .EndElement="<!---TopNewsList07:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = ""
    For I=1 To 8
        If Rs.Eof And Rs.Bof Then
            Exit For
        End If
        strHtml = strHtml & "·[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!---TopNewsList08:start-->"
        .EndElement="<!---TopNewsList08:end-->"
        .Value=strHtml
        .ReplaceTemplate()        
    End With

    Rs.Close
    Set Rs=Nothing

    TClass.Save()
    Set TClass=Nothing
End Function
%>