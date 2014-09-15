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
'//��ҳ��
'//����ͼƬ������ҳ����
'////////////////////////////////////////////////////////////////////

Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If

If Not SysAdmin.UpdatePage Then
    Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
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
        UpdateOk()
End Select

'//���³ɹ���ʾ
Function UpdateOk()
    Response.Write("<script>alert(""<�����ɹ�>\nҳ����³ɹ�"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End Function

'//��ͼ����ͼ�Ƽ�
Sub Update01()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../pcenter/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 3 Title,FilePath,SmallImg,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',22,')<>0 Order By Id DESC"
	Set Rs=Conn.ExeCute(Sql)
	While Not Rs.Eof
		strHtml=strHtml&"<td><a href="""&Rs("FilePath")&""" target=""_blank""><img src="""&Rs("SmallImg")&""" width=""160"" height=""120"" border=""1"" style=""border-color:#000000""></a></td>" & vbCrLf
		Rs.MoveNext
	Wend
	strHtml="<table width=""100%""  border=""0"" cellspacing=""0"" height=""120"" cellpadding=""2""><tr align=""center"" valign=""middle"">" & strHtml & "</tr></table>"
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--CoolList:start-->"
        .EndElement="<!--CoolList:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Sub

'//�б���ͼ�Ƽ�
Function Update02()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../pcenter/default.htm")
    Dim TClass,strHtml,I
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 16 Title,FilePath,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',26,')<>0 Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)

    TClass.OpenTemplate(TemplateFilePath)

    strHtml = ""
    For I=1 To 8
        If Rs.Eof Then
            Exit For
        End If
        strHtml = strHtml & "��[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!--CoolList01:start-->"
        .EndElement="<!--CoolList01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = ""
    For I=1 To 8
        If Rs.Eof Then
            Exit For
        End If
        strHtml = strHtml & "��[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!--CoolList02:start-->"
        .EndElement="<!--CoolList02:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    Rs.Close
    Set Rs=Nothing

    TClass.Save()
    Set TClass=Nothing
End Function

'//�б���Ϸ��ͼ
Function Update03()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../pcenter/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 16 Title,FilePath,AddTime From view_NewsInfo Where Class In (40" & AllChildClass(40) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        strHtml=strHtml&"��[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing

    With TClass
        .OpenTemplate(TemplateFilePath)
        .StartElement="<!--GamePicList01:start-->"
        .EndElement="<!--GamePicList01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Function

'//������ֽ
Sub Update04()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../pcenter/default.htm")
    Dim TClass,strHtml
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 4 Title,FilePath,SmallImg,AddTime From view_NewsInfo Where Instr(','+Speciality+',',',24,')<>0 Order By Id DESC"
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
        .StartElement="<!--BestPicture:start-->"
        .EndElement="<!--BestPicture:end-->"
        .Value=strHtml
        .ReplaceTemplate()
        .Save()
    End With
    Set TClass=Nothing
End Sub

'//�б�������ֽ����
Function Update05()
    Dim TemplateFilePath
        TemplateFilePath=Server.MapPath("../../../pcenter/default.htm")
    Dim TClass,strHtml,I
    Set TClass=New Tkl_TemplateClass
        strHtml=""
    Dim Rs,Sql
        Sql="Select Top 16 Title,FilePath,AddTime From view_NewsInfo Where Class In (41" & AllChildClass(41) & ") Order By Id DESC"
    Set Rs=Conn.ExeCute(Sql)

    TClass.OpenTemplate(TemplateFilePath)

    strHtml = ""
    For I=1 To 8
        If Rs.Eof Then
            Exit For
        End If
        strHtml = strHtml & "��[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!--PictureList01:start-->"
        .EndElement="<!--PictureList01:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    strHtml = ""
    For I=1 To 8
        If Rs.Eof Then
            Exit For
        End If
        strHtml = strHtml & "��[" & StrClass.FormatMyDate(Rs("AddTime"),"{m}/{d}") & "]<a href=""" & Rs("FilePath") & """ target=""_blank"">" & Rs("Title") & "</a><br>" & vbCrLf
        Rs.MoveNext
    Next

    With TClass
        .StartElement="<!--PictureList02:start-->"
        .EndElement="<!--PictureList02:end-->"
        .Value=strHtml
        .ReplaceTemplate()
    End With

    Rs.Close
    Set Rs=Nothing

    TClass.Save()
    Set TClass=Nothing
End Function
%>