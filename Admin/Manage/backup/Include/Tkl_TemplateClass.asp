<%
'////////////////////////////////////////////////////////////
'//��ҳģ��������(TklMilk Boy9732@msn.com)
'//��;����Ҫ���ھ�̬��Դ��ҳ������..����Ҳ����������ϵͳ�ľ�̬ҳ�������
'//����ʵ����
'// Dim TClass
'// Set TClass=New Tkl_TemplateClass
'//     TClass.OpenTemplate("e:/t.htm")
'//     TClass.StartElement="<!--��Դ�����ʶ-��ʼ-->"
'//     TClass.EndElement="<!--��Դ�����ʶ-����-->"
'//     TClass.Value="������Դ������滻����"
'//     TClass.ReplaceTemplate()
'// 
'//     TClass.StartElement="<!--��Դ����-��ʼ-->"
'//     TClass.EndElement="<!--��Դ����-����-->"
'//     TClass.Value="�ҽ��滻����Դ������"
'//     TClass.ReplaceTemplate()
'// 
'//     TClass.Save()
'//     TClass.SaveAs("e:/t1.htm")
'// Set TClass=Nothing
'////////////////////////////////////////////////////////////

Class Tkl_TemplateClass
    Public FilePath                             'ģ���ļ�
    Private Template                            'ģ���ݴ����
    Public StartElement                         'Ԫ�ؿ�ʼ��ǩ
    Public EndElement                           'Ԫ�ؽ�����ǩ
    Public Value                                '��������

    Private Fso
    Private Fle
    Private regEx
    Private FileState                            '�ļ�״̬

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

    '//����ģ���ļ�·��
    Public Function OpenTemplate(mFilePath)
        Set Fle=Fso.OpenTextFile(mFilePath,1)
        Template=Fle.ReadAll
        Fle.Close
        FileState=True
        FilePath=mFilePath
    End Function

    '//�滻ģ��Ԫ��,Ԫ�ر�ǩһ���ʽΪ����<!-Ԫ�ر�ǩ-��ʼ--><!--Ԫ�ر�ǩ-����-->��,�㵱ȻҲ��������
    '//���ر�ǩ�����ִ�Сд
    Public Function ReplaceTemplate()
        If (Not FileState) Or Template="" Or StartElement="" Or EndElement="" Then
            Exit Function
        End If
        Dim strPatrn
            strpatrn=FilterStr(StartElement) & "[\S\s]*?" & FilterStr(EndElement)
        regEx.Pattern = strPatrn
        Template=regEx.Replace(Template,StartElement & vbCrLf & Value & vbCrLf & EndElement)
    End Function

    '//�����µ�ģ������
    Public Function Save()
        If (Not FileState) Then
            Exit Function
        End If
        Set Fle=Fso.OpenTextFile(FilePath,2)
        Fle.Write Template
        Fle.Close
    End Function

    '//���ģ������
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
