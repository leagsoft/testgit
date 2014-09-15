<%
'////////////////////////////////////////////////////////////
'//网页模版生成类(TklMilk Boy9732@msn.com)
'//用途：主要用于静态资源面页的生成..本类也适用于其它系统的静态页面的生成
'//调用实例：
'// Dim TClass
'// Set TClass=New Tkl_TemplateClass
'//     TClass.OpenTemplate("e:/t.htm")
'//     TClass.StartElement="<!--资源标题标识-开始-->"
'//     TClass.EndElement="<!--资源标题标识-结束-->"
'//     TClass.Value="这是资源标题的替换内容"
'//     TClass.ReplaceTemplate()
'// 
'//     TClass.StartElement="<!--资源内容-开始-->"
'//     TClass.EndElement="<!--资源内容-结束-->"
'//     TClass.Value="我将替换成资源的内容"
'//     TClass.ReplaceTemplate()
'// 
'//     TClass.Save()
'//     TClass.SaveAs("e:/t1.htm")
'// Set TClass=Nothing
'////////////////////////////////////////////////////////////

Class Tkl_TemplateClass
    Public FilePath                             '模板文件
    Private Template                            '模板暂存变量
    Public StartElement                         '元素开始标签
    Public EndElement                           '元素结束标签
    Public Value                                '插入内容

    Private Fso
    Private Fle
    Private regEx
    Private FileState                            '文件状态

    Private Sub Class_Initialize
        FileState=False
        FilePath=""
        Template=""
        StartElement=""
        EndElement=""
        Set regEx = New RegExp
            With regEx
                .Multiline = True
                .IgnoreCase = True
                .Global = True
            End With
        Set Fso = Server.CreateObject("Scripting.FileSystemObject")
    End Sub

    Private Sub class_Terminate
        FilePath=""
        Template=""
        StartElement=""
        EndElement=""
        Set regEx=Nothing
        Set Fle=Nothing
        Set Fso=Nothing
    End sub

    Private Function FilterStr(str)
        FilterStr=str
        If str="" Or IsNull(FilterStr) Then
            FilterStr=""
        Else
            FilterStr=Replace(FilterStr,"\","\\")
            FilterStr=Replace(FilterStr,"(","\(")
            FilterStr=Replace(FilterStr,")","\)")
            FilterStr=Replace(FilterStr,"*","\*")
            FilterStr=Replace(FilterStr,"?","\?")
            FilterStr=Replace(FilterStr,"{","\{")
            FilterStr=Replace(FilterStr,"}","\}")
            FilterStr=Replace(FilterStr,".","\.")
            FilterStr=Replace(FilterStr,"+","\+")
            FilterStr=Replace(FilterStr,"[","\[")
            FilterStr=Replace(FilterStr,"]","\]")
        End If
    End Function

    '//设置模板文件路径
    Public Function OpenTemplate(mFilePath)
        Set Fle=Fso.OpenTextFile(mFilePath,1)
        Template=Fle.ReadAll
        Fle.Close
        FileState=True
        FilePath=mFilePath
    End Function

    '//规换模版元素,元素标签一般格式为：＂<!-元素标签-开始--><!--元素标签-结束-->＂,你当然也可以自义
    '//无素标签不区分大小写
    Public Function ReplaceTemplate()
        If (Not FileState) Or Template="" Or StartElement="" Or EndElement="" Then
            Exit Function
        End If
        Dim strPatrn
            strpatrn=FilterStr(StartElement) & "[\S\s]*?" & FilterStr(EndElement)
        regEx.Pattern = strPatrn
        Template=regEx.Replace(Template,StartElement & vbCrLf & Value & vbCrLf & EndElement)
    End Function

    '//保存新的模板内容
    Public Function Save()
        If (Not FileState) Then
            Exit Function
        End If
        Set Fle=Fso.OpenTextFile(FilePath,2)
        Fle.Write Template
        Fle.Close
    End Function

    '//另存模板内容
    Public Function SaveAs(mFilePath)
        If (Not FileState) Then
            Exit Function
        End If
        Set Fle=Fso.OpenTextFile(mFilePath,2,1)
        Fle.Write Template
        Fle.Close
    End Function
End Class
%>
