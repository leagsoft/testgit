<%
'////////////////////////////////////////////////////////////
'                    ������Ա�����ࣨ�����ݿ⣩
'�������:
'���ʱ��:2003-05-16
'////////////////////////////////////////////////////////////


'-----------------------------------------------------------
'                        ��Ա��Ϣ��
'-----------------------------------------------------------
Class UserInfo_Class
    Private arrInfo_User(6)
    Private Field_SplitStr

    Public Property Get InfoItme_User(byval val)
        If val <= UBound(arrInfo_User) Then
            InfoItme_User = arrInfo_User(val)
        Else
            InfoItme_User = ""
        End If
    End Property

    Public Property Let InfoItme_User(byval valIndex,byval valValue)
        If valIndex <= UBound(arrInfo_User) Then
            arrInfo_User(valIndex) = valValue
        End If
    End Property

    Public Property Get Id()
        Id = arrInfo_User(0)
    End Property

    Public Property Let Id(byval val)
        arrInfo_User(0) = val
    End Property

    Public Property Get Name()
        Name = arrInfo_User(1)
    End Property

    Public Property Let Name(byval val)
        arrInfo_User(1) = val
    End Property

    Public Property Get IP()
        IP = arrInfo_User(2)
    End Property

    Public Property Let IP(byval val)
        arrInfo_User(2) = val
    End Property

    Public Property Get AddTime()
        AddTime = arrInfo_User(3)
    End Property

    Public Property Let AddTime(byval val)
        arrInfo_User(3) = val
    End Property

    Public Property Get UpTime()
        UpTime = arrInfo_User(4)
    End Property

    Public Property Let UpTime(byval val)
        arrInfo_User(4) = val
    End Property

    Public Property Get Remark()
        Remark = arrInfo_User(5)
    End Property

    Public Property Let Remark(byval val)
        arrInfo_User(5) = val
    End Property

    Public Property Get NickName()
        NickName = arrInfo_User(6)
    End Property

    Public Property Let NickName(byval val)
        arrInfo_User(6) = val
    End Property


    Private Sub Class_Initialize
        Dim I
        For I = 0 To UBound(arrInfo_User)
            arrInfo_User(I) = ""
        Next
        '
        Field_SplitStr = "{-}"
    End Sub

    Public Function ToString()
        Dim Temp_Str
            Temp_Str = ""
        Dim I
        For I=0 To UBound(arrInfo_User)
            If Temp_Str = "" Then
                Temp_Str = arrInfo_User(I)
            Else
                Temp_Str = Temp_Str & Field_SplitStr & arrInfo_User(I)
            End If
        Next
        ToString = Temp_Str
    End Function

    Public Property Get SplitStr()
        SplitStr = Field_SplitStr
    End Property
    
    ' ���û���Ϣ�ַ���ת��Ϊ��׼�û���Ϣ��
    Public Function strToClass(byval val)
        If val="" Then
            Set strToClass = Nothing
            Exit Function
        End If
        Dim Temp_User
        Set Temp_User = New UserInfo_Class
        Dim Temp_UserInfo_Item
            Temp_UserInfo_Item = Split(val,SplitStr(),-1,1)
        Dim I
        For I = 0 To UBound(Temp_UserInfo_Item)
            Temp_User.InfoItme_User(I) = Temp_UserInfo_Item(I)
        Next

        Set strToClass = Temp_User
    End Function

    Public Function UpdateTime()
        UpTime = Now
    End Function

End Class

'-----------------------------------------------------------
'                         ��Ա��Ϣ�ռ�����
'-----------------------------------------------------------
Class UserInfo_Collection_Class
    
    Private mTimeOut                                                '��ʱʱ�䣨���ӣ�
    Private User_SplitStr                                           '����Ա��ķָ���
    Private ApplicationSaveSize                                    'Application("UserInfo_Collection")�İ�ȫ��С(�ֽ�)

    Private Sub Class_Initialize
        ApplicationSaveSize=1024*1024*1                             '1M
        mTimeOut = Session.Timeout                                  '��ʼ��Ĭ�ϳ�ʱʱ��
        User_SplitStr = "{+}"
        If Len(Application("UserInfo_Collection"))>ApplicationSaveSize Then
            Clear()
        End If
        Refh()
    End Sub

    Public Property Get TimeOut()                                   'ȡ�õ�ǰ��ʱʱ��(��)
        TimeOut = mTimeOut
    End Property

    Public Property Let TimeOut(byval val)                          'val:����
        mTimeOut = val
        Session.Timeout = val
    End Property

    '������Ա��������Name��
    Public Function Find(byval val)
        Dim User_temp
        Set User_temp = GetUser(val)
        Find=Not (User_temp.Name = "" And User_temp.Id = "" And User_temp.AddTime ="")
    End Function

    '�������Ա(ͬ���򸴸���Ϣ)
    Public Function Add(byval val)
        Dim User_temp
        Set User_temp = New UserInfo_Class
        Set User_temp = val
        If Not Find(User_temp.Name) Then
            ' �����ڴ��û������
            Dim AppVal
                Application.Lock
                AppVal = Application("UserInfo_Collection")
                Application.UnLock
            If AppVal = "" Then
                AppVal = User_temp.ToString()
            Else
                AppVal = AppVal & User_SplitStr & User_temp.ToString()
            End If
            Application.Lock
            Application("UserInfo_Collection") = AppVal
            Application.UnLock
        Else
            ' ��������¸���Ա��Ϣ
            Dim List_User
                List_User = Split(ToString(),User_SplitStr,-1,1)
            Dim Temp_UserInfo
            Set Temp_UserInfo = New UserInfo_Class
            Dim I
            For I = 0 To UBound(List_User)
                Set Temp_UserInfo = Temp_UserInfo.strToClass(List_User(I))
                If Temp_UserInfo.Name = User_temp.Name Then
                    List_User(I) = User_temp.ToString()
                    Exit For
                End If
            Next
            '������������û���Ϣ���ռ�����
            Dim Temp_AppVal
                Temp_AppVal = ""
            For I = 0 To UBound(List_User)
                If List_User(I)<>"" Then
                    If Temp_AppVal = "" Then
                        Temp_AppVal = List_User(I)
                    Else
                        Temp_AppVal = Temp_AppVal & SplitStr() & List_User(I)
                    End If
                End If
            Next
            Application.Lock
            Application("UserInfo_Collection") = Temp_AppVal
            Application.UnLock
        End If
    End Function

    '������Ա�ռ���������Ϣ
    Public Function ToString()
        Application.Lock
        ToString = Application("UserInfo_Collection")
        Application.UnLock
    End Function

    '�����Ա�ռ���
    Public Function Clear()
        Application.Lock
        Application("UserInfo_Collection") = ""
        Application.UnLock
    End Function

    '���ص�ǰ������Ա��Ŀ
    Public Property Get Count()
        If ToString()="" Then
            Count = 0
        Else
            Count = UBound(Split(ToString(),User_SplitStr,-1,1)) + 1
        End If
    End Property

    '����û���Ϣ,(val Ϊ�û����ƻ�����û�λ��<��ֵ��>)
    '
    Public Function GetUser(byval val)
        Dim Finded
            Finded = False
        Dim Name_UserFind
            Name_UserFind = val
        Dim List_User
            List_User = Split(ToString(),User_SplitStr,-1,1)
        Dim I
        Dim Temp_UserInfo
        Set Temp_UserInfo = New UserInfo_Class
        If Name_UserFind <> "" Then
            If ISNumeric(val) Then                                      '��������ȡ���û�
                val = CInt(val)
                If val <= (UBound(List_User)+1) Then
                    Finded = True
                    Set GetUser = Temp_UserInfo.strToClass(List_User(val-1))
                End If
            Else                                                        '�����û��������û�
                For I = 0 To UBound(List_User)
                    Set Temp_UserInfo = Temp_UserInfo.strToClass(List_User(I))
                    If UCase(Temp_UserInfo.Name) = UCase(val) Then
                        Set GetUser = Temp_UserInfo
                        Finded = True
                        Exit For
                    End If
                Next
            
            End If
        End If

        '��δ�ҵ��û����򽫷���һ��������ϢΪ�յ��û���
        If Not Finded Then
            Set GetUser = New UserInfo_Class
        End If
    End Function

    'ɾ��ָ���û�(�������ƻ�λ��)
    Public Function Remove(byval val)
'        On Error Resume Next
        Dim Deleted
            Deleted = False
        Dim List_User
            List_User = Split(ToString(),User_SplitStr,-1,1)
        Dim I
        Dim Temp_UserInfo
        Set Temp_UserInfo = New UserInfo_Class
        For I = 0 To UBound(List_User)
            Set Temp_UserInfo = Temp_UserInfo.strToClass(List_User(I))
            If Temp_UserInfo.Name = val Then
                List_User(I) = ""
                Deleted = True
                Exit For
            End If
        Next


        '������������û���Ϣ���ռ�����
        Dim Temp_AppVal
            Temp_AppVal = ""
        For I = 0 To UBound(List_User)
            If List_User(I)<>"" Then
                If Temp_AppVal = "" Then
                    Temp_AppVal = List_User(I)
                Else
                    Temp_AppVal = Temp_AppVal & SplitStr() & List_User(I)
                End If
            End If
        Next
        Application.Lock
        Application("UserInfo_Collection") = Temp_AppVal
        Application.UnLock
        Remove = Deleted
    End Function

    Public Property Get SplitStr()
        SplitStr = User_SplitStr
    End Property

    '��ˢ���û���Ϣ
    Public Function Refh()
        Dim NowTime
            NowTime = Now()
        Dim List_User
            List_User = Split(ToString(),User_SplitStr,-1,1)
        Dim Temp_UserInfo
        Set Temp_UserInfo = New UserInfo_Class
        Dim I
        For I = 0 To UBound(List_User)
            Set Temp_UserInfo = Temp_UserInfo.strToClass(List_User(I))
            If DateDiff("s",Temp_UserInfo.UpTime,NowTime) > (mTimeOut*60) Then
                Remove(Temp_UserInfo.Name)
            End If
        Next
    End Function

End Class
%>