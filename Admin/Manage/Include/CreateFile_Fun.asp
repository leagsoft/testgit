<%
'////////////////////////////////////////
'//函数：将指定资源内容套入模板当中
'//参数：资源Id
'//返回：Bool,是否成功
Function UsedTemplate_CreateFile(NewsId)
'    on error resume next
    Dim Sql
    '//取得资源信息
    Dim Rs1
    Sql="Select * From view_NewsInfo2 Where Id=" & CLng(NewsId)
    Set Rs1=Conn.ExeCute(Sql)
    If Rs1.Eof And Rs1.Bof Then
        Rs1.Close
        Set Rs1=Nothing
        UsedTemplate_CreateFile=false
        Exit Function
    End If

    '//若是链接跳转型资料直接无需生成
    If Trim(Rs1("Url"))="" Or IsNull(Rs1("Url")) Then
        '如果模板为空，则不创建
        Dim Template
        '缓冲区中资源模板所属的类别Id
        If Buffer_WhenCreatingFile And Session("buffer_NewsTemplate_ClassId")=Rs1("Class") And Session("buffer_NewsTemplate")<>""Then
            Template=Session("buffer_NewsTemplate")
        Else
            Template=GetTemplate(Rs1("Class"))
        End If

        If Trim(Template)="" Or ISNULL(Template) Then
            Rs1.Close
            Set Rs1=Nothing
            UsedTemplate_CreateFile=false
            Exit Function
        End If

        Dim charClass
        set charClass = new Tkl_StringClass
        '替换模板内容

        Dim str_patrn
        str_patrn="<title>.*?</title>"
        Template=charClass.ReplaceTest(str_patrn,Template,"<title>" & charClass.GetTextFromHtml(Rs1("Title")) & " - "&Def_MySiteTitle&"</title>")
        str_patrn="\$Id\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("Id")))
        str_patrn="\$Title\$"	
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("Title")))
        str_patrn="\$Author\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("AuthorTitle")))
        str_patrn="\$From\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("FromContent")))
        str_patrn="\$ClassTitle\$"
        Template=charClass.ReplaceTest(str_patrn,Template,""&Rs1("ClassTitle"))
        str_patrn="\$ClassTitle2\$"
        Template=charClass.ReplaceTest(str_patrn,Template,""&Rs1("ClassTitle2"))
        str_patrn="\$ClassUrl\$"
        Template=charClass.ReplaceTest(str_patrn,Template,""&Rs1("ClassUrl"))
        str_patrn="\$KeyWord\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("KeyWord")))
        str_patrn="\$Editor\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("EditorTitle")))
        str_patrn="\$SmallImg\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("SmallImg")))
        str_patrn="\$BigImg\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("BigImg")))
        str_patrn="\$ShortContent\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("ShortContent")))
        str_patrn="\$AddTime\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("AddTime")))
        str_patrn="\$UpTime\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&Rs1("UpTime")))
        str_patrn="\$Count\$"
        Template=charClass.ReplaceTest(str_patrn,Template,"<script src="""&TsysRootPath&"/Count.asp?Id=" & Rs1("Id") & """></script>")
        str_patrn="\$CommentCount\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr("<script src="""&TsysRootPath&"/Comment/CommenCount.asp?Id="& Rs1("Id") & """></script>"))
        str_patrn="\$Comment\$"
        Template=charClass.ReplaceTest(str_patrn,Template,TsysRootPath&"/Comment/Comment_List.asp?Id="& Rs1("Id") &"&ResTitle="&Server.UrlEncode(charClass.GetTextFromHtml(Rs1("Title"))))
        str_patrn="\$ConnectNewsList\$"
        Template=charClass.ReplaceTest(str_patrn,Template,CStr(""&GetConnectNewsList(Rs1("KeyWord"),Rs1("Id"))))

        '//处理分页-开始
        Dim temp_Template
            temp_Template=Template
        Dim arrContent
            arrContent=Split(Rs1("Content"),"<HR sysPageSplitFlag>",-1,1)
        str_patrn="\$Content\$"

        Dim I,J
        Dim Fso
        Set Fso = Server.CreateObject(FsoObjectStr)
        Dim Fle
        Dim FilePath,FileLocalPath
        Dim strPageList

        For I=0 To UBound(arrContent)
            If I=0 Then
                '生成文件存放路径
                FilePath=CreateFileSaveToPath(CInt(NewsId),Rs1("AddTime"),Rs1("Directory"))
                FileLocalPath=CreateFileLocalPath(CInt(NewsId),Rs1("AddTime"),Rs1("Directory"))
            Else
                FilePath=CreateFileSaveToPath(CInt(NewsId),Rs1("AddTime"),Rs1("Directory"))
                FilePath=Left(FilePath,(Len(FilePath)-Len(ExNameOfNewsFile)))&"_"& I & ExNameOfNewsFile
            End If

            strPageList=""
            If UBound(arrContent)>=1 Then
                For J=0 To UBound(arrContent)
                    If J=I Then
                        strPageList=strPageList&"<b>["&J+1&"]</b>&nbsp;"
                    Else                    
                        If J=0 Then
                        '第一页
                            strPageList=strPageList&"<a href='" & NewsId & ExNameOfNewsFile & "'>["&J+1&"]</a>&nbsp;"
                        Else
                            strPageList=strPageList&"<a href='" & NewsId & "_"&J & ExNameOfNewsFile &"'>["&J+1&"]</a>&nbsp;"
                        End If
                    End If
                Next
            End If

            '生成上一页/下一页
            If UBound(arrContent)>0 Then
                If 0<I Then
                    If I=1 Then
                        strPageList="<a href='" & NewsId & ExNameOfNewsFile & "'>上一页</a>&nbsp;" & strPageList
                    Else
                        strPageList="<a href='" & NewsId & "_"& I-1 & ExNameOfNewsFile &"'>上一页</a>&nbsp;" & strPageList
                    End If
                End If
                If I<UBound(arrContent) Then
                    strPageList=strPageList & "<a href='" & NewsId & "_"& I+1 & ExNameOfNewsFile &"'>下一页</a>"
                End If
            End If
            Template=charClass.ReplaceTest(str_patrn,temp_Template,arrContent(I)&"<p><center>"&strPageList&"</center></p>")
            '如果文件不存在，则创建，存在则复盖
            Set Fle = Fso.OpenTextFile(FilePath,2,true)
            Fle.Write Template
            Fle.Close
        Next
        '//处理分页-结束
    Else
        FileLocalPath = Rs1("Url")
    End If
    Set Fle=Nothing
    Set Fso=Nothing
    Rs1.Close
    Set Rs1=Nothing

    If err.Number<>0 Then
        UsedTemplate_CreateFile=false
    Else
        Sql="UPDATE News Set Created=1,FilePath='"&FileLocalPath&"' Where Id=" & NewsId
        Conn.ExeCute(Sql)
        UsedTemplate_CreateFile=true
    End If
End Function

'////////////////////////////////////////
'//函数：删除指定资源的静态文件(包含分页)
'//参数：资源路径,资源Id
Function DeleteNewsFile(FilePath,Id)
    '若是链接跳转型资料直接退出
    If Not IsLocalFilePath(FilePath) Then
        Exit Function
    End If

    Dim Fso
    Set Fso = Server.CreateObject(FsoObjectStr)
    FilePath = Server.MapPath(FilePath)
    If Fso.FileExists(FilePath) Then
        Fso.DeleteFile(FilePath)
        '删除所有分页
        Dim I
            I=0
        Dim SplitPage_FilePath
        While(I<>-1)
            I=I+1
            SplitPage_FilePath=Replace(FilePath,Id&".",Id&"_"&I&".")
            If Fso.FileExists(SplitPage_FilePath) Then
                Fso.DeleteFile(SplitPage_FilePath)
            Else
                I=-1
            End If
        Wend
    End If
End Function

'////////////////////////////////////////
'//文件路径是否为本地路径
Function IsLocalFilePath(FilePath)
    If Trim(FilePath)="" Or IsNull(FilePath) Then
        IsLocalFilePath = False
        Exit Function
    End If

    Dim regEx
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Pattern = "^/"
    '若是链接跳转型资料直接退出
    IsLocalFilePath = regEx.Test(FilePath)
    Set regEx = Nothing
End Function

'////////////////////////////////////////
'//删除文件
Function DelFile(fPath)
    Dim Fso
    Set Fso = Server.CreateObject(FsoObjectStr)
    If Fso.FileExists(fPath) Then
        Fso.DeleteFile(fPath)
    End If
    Set Fso=Nothing
End Function

'////////////////////////////////////////
'//函数：生成文件逻辑存放路径（不创建目录）
'//参数：类别Id,资源Id,资源添加时间
'//返回：字符串,格式：12/2003040506/342.htm
Function CreateFileLocalPath(NewsId,AddTime,Directory)
    Dim tPath
    tPath = Directory & "/" & Create_id(AddTime)&"/"&NewsId & ExNameOfNewsFile
    CreateFileLocalPath=tPath 
End Function

'////////////////////////////////////////
'//函数：生成文件物理存放路径(创建相关目录)
'//参数：类别Id,资源Id,资源添加时间,是否使用指定目录,指定目录地址
'//返回：资源最终生成路径,如：/12/2003040504/2322.htm
Function CreateFileSaveToPath(NewsId,AddTime,Directory)
    Dim Fso
    Set Fso = Server.CreateObject(FsoObjectStr)
    Dim tPath

    tPath = Server.MapPath(Directory)

    If Not Fso.FolderExists(tPath) Then
        Fso.CreateFolder(tPath)
    End If

    tPath=tPath & "/"&Create_id(AddTime)
    If Not Fso.FolderExists(tPath) Then
        Fso.CreateFolder(tPath)
    End If

    Set Fso=Nothing
    CreateFileSaveToPath=tPath & "/"&NewsId & ExNameOfNewsFile
End Function

'////////////////////////////////////////
'//函数：创值ID值
'//返回：Year+Month+Day
Function Create_id(cTime)
    Create_id=Year(cTime) & Right("00"&Month(cTime),2) & Right("00"&Day(cTime),2)
End Function

'////////////////////////////////////////
'//函数：取得相关资源列表
'//参数：资源关键词,资源Id
'//返回：Html字符串
Function GetConnectNewsList(kWord,NewsId)
    Dim Result
        Result=""
    If kWord="" Or IsNULL(kWord) Then
        GetConnectNewsList=Result
        Exit Function
    End If
    Dim arr,I
        arr=Split(kWord,",",NewsKeyWordListNum,1)
    Dim tSql
        tSql=""

    For I=0 To UBound(arr)
        If tSql<>"" THen
            tSql=tSql & " OR Title Like '%"&arr(I)&"%' OR KeyWord Like '%"&arr(I)&"%'"
        Else
            tSql=tSql & " Title Like '%"&arr(I)&"%' OR KeyWord Like  '%"&arr(I)&"%'"
        End If
    Next

    If tSql<>"" Then
        tSql=" Where (" & tSql &") And Id<>"&NewsId &_
             " Order By Id DESC"
    End If

    Dim Rs,Sql
        Sql="Select Top " & RelateNewsNumber & " Id,Title,Class,FilePath,AddTime From view_NewsInfo" & tSql
    Set Rs=Conn.ExeCute(Sql)

    While Not Rs.Eof
        Result=Result & "<li><a href=""" & Rs("FilePath") & """>"&Rs("Title")&"</a>"
        Result=Result & " ["&FormatDateTime(Rs("AddTime"),2)&"]" & "</li>"
        Rs.MoveNext
    Wend

    Rs.Close
    Set Rs=Nothing
    GetConnectNewsList=Result
End Function

'//////////////////////////////////////
'//函数：取得类别的模板信息
'//参数：资源类别Id
'//返回：模板内容
Function GetTemplate(ClassId)
    Dim Sql,Rs2
    GetTemplate=""
    Sql="Select Top 1 CL.Id,CL.Title,CL.UpTime,NT.Id As TemplateId,NT.Content As Template From ClassList CL LEFT JOIN News_Template NT ON CL.Template=NT.Id Where CL.Id = " & ClassId
    Set Rs2=Conn.ExeCute(Sql)
    If Not(Rs2.Eof And Rs2.Bof) Then
        GetTemplate=Rs2("Template")
        If Buffer_WhenCreatingFile Then
            '将当前使用的模板信息（模板内容及模板所属的类别）缓冲
            Session("buffer_NewsTemplate")=GetTemplate
            Session("buffer_NewsTemplate_ClassId")=ClassId
        Else
            Session("buffer_NewsTemplate")=""
        End If
    End If
    Rs2.Close
    Set Rs2=Nothing
End Function
%>