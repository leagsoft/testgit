<%
'////////////////////////////////////////////////////////////
'//��Դϵͳ����Ա��Ϣ��(TklMilk Boy9732@msn.com)
'//��;���Թ���Ա����Ϣ��ȡ,Ȩ���жϽ��й淶��ͳһ����
'//����ʵ����
'//Dim SysAdmin
'//Set SysAdmin=New SYSProedom_Class
'//Set SysAdmin=Nothing
'//���ԣ�
'//------------������-------------��������------------------���-----
'//         .Logined             Bool                ����Ա�Ƿ��¼
'//         .AdminTitle          String              ����Ա�ʻ�����
'//         .AdminNickName       String              ����Ա�ǳ�
'//         .AdminRoleTitle      String              ����Ա������ɫ����
'//         �����������������ע��...
'////////////////////////////////////////////////////////////

Class SYSProedom_Class

    Private PopedomList
	Private ClassPopedomList

    Private Podm_ChagePWD
    Private Podm_ChageClass
    Private Podm_ChageFrom
    Private Podm_ChageAuthor
    Private Podm_ChageTopAndCpr
    Private Podm_ChageNews
    Private Podm_ChageDustbin
    Private Podm_ManageFiles
    Private Podm_ConfigSite
    Private Podm_ChangeAdminList
    Private Podm_ChangeRole
    Private Podm_ChangeCommentList
    Private Podm_MoveNews
    Private Podm_Speciality
    Private Podm_CommentList
    Private Podm_NewsTemplate
    Private Podm_SysConfig
    Private Podm_ManageDataBase
    Private Podm_UpdatePage
    Private Podm_InsertSYS
    Private Podm_CreateNewsFile

'////////////////////////////////////////
'  �������������Ȩ�ޱ�������
'////////////////////////////////////////

    Private Sub Class_Initialize
        '��ʼ��Ȩ����ĿΨһ��ʶ
        Podm_ChagePWD=1
        Podm_ChageClass=2
        Podm_ChageFrom=3
        Podm_ChageAuthor=4
        Podm_ChageTopAndCpr=5
        Podm_ChageNews=6
        Podm_ManageFiles=7
        Podm_ChageDustbin=8
        Podm_ConfigSite=10
        Podm_ChangeAdminList=11
        Podm_ChangeRole=12
        Podm_ChangeCommentList=13
        Podm_MoveNews=14
        Podm_Speciality=15
        Podm_CommentList=16
        Podm_NewsTemplate=17
        Podm_SysConfig=18
        Podm_ManageDataBase=19
        Podm_UpdatePage=20
        Podm_InsertSYS=21
        Podm_CreateNewsFile=22

'////////////////////////////////////////
'  �������ʼ������Ȩ�޳���
'////////////////////////////////////////

        '��ʼ��Ȩ���б�
        PopedomList=Podm_ChagePWD&",�޸�����,"&_
                    Podm_ChageClass&",�޸���Ŀ,"&_
                    Podm_ChageFrom&",�޸���Դ,"&_
                    Podm_ChageAuthor&",�޸�����,"&_
                    Podm_ChageNews&",����������Դ,"&_
                    Podm_ChageDustbin&",����վ����,"&_
                    Podm_ChangeCommentList&",���۹���,"&_
                    Podm_ManageFiles&",�ļ�ϵͳ����,"&_
                    Podm_UpdatePage&",վ��ҳ�����,"&_
                    Podm_InsertSYS&",վ�������滻,"&_
                    Podm_CreateNewsFile&",������Դ,"&_
                    Podm_MoveNews&",�ƶ���Դ,"&_
                    Podm_Speciality&",��Դ����,"&_
                    Podm_CommentList&",���۹���,"&_
                    Podm_NewsTemplate&",��Դģ��,"&_
                    Podm_ChangeAdminList&",�û�����,"&_
                    Podm_ChangeRole&",��ɫ����,"&_
                    Podm_SysConfig&",ϵͳ����,"&_
                    Podm_ManageDataBase&",���ݿ����,"&_
                    ""

    End Sub

    Public Property Get defPopedomList()
        defPopedomList=PopedomList
    End Property

    'Ĭ�ϵ�ϵͳ��������Ա������,���û����ɱ�ɾ�����޸�
    Public Property Get defAdminUserTitle()
        defAdminUserTitle="Admin"
    End Property

    '����AdminRoleTitle�Ľ�ɫ���ɱ�ɾ�������޸�
    Public Property Get defAdminRoleTitle()
        defAdminRoleTitle="��������Ա"
    End Property

    '��ĿȨ������,��
    Public Property Get defClassPopedomType_Low()
        defClassPopedomType_Low=0
    End Property
    '��ĿȨ������,��
    Public Property Get defClassPopedomType_Mid()
        defClassPopedomType_Mid=1
    End Property
    '��ĿȨ������,��
    Public Property Get defClassPopedomType_Hig()
        defClassPopedomType_Hig=2
    End Property

    '����Ա�Ƿ��¼,����bool
    Public Property Get Logined()
        If Session("AdminLogined")="TRUE" And Trim(Session("AdminTitle"))<>"" Then
            Logined=true
        Else
            Logined=false
        End If
    End Property
    '����Ա�˳�
    Public Sub LogOut()
        Session.Abandon
    End Sub

    '////////////////////////////////////////////////////////
    '//�ж�mItem�Ƿ���stritemList�б��У����б���splitStrΪ�ָ���
    '����3 �Ƿ� 1,2,3,4,5
    Private Function ItemInList(mItem,strItemList,splitStr)
        Dim I
        Dim ItemList
        ItemInList=false
        If splitStr="" Then
            splitStr=","
        End If
        mItem=Trim(CStr(mItem))
        If mItem<>"" And strItemList<>"" then
            ItemList=Split(strItemList,splitStr,-1,1)
            For I=0 To UBound(ItemList)
                If Trim(CStr(ItemList(I)))=mItem Then
                    ItemInList=True
                    Exit For
                End If
            Next
        End If
    End Function

    '////////////////////////////////////////////////////////
    '//���ָ������ĵĲ���Ȩֵ,
    '//���أ������򷵻�Ȩֵ;��Ȩ���򷵻�-1
    Public Function EnoughClassPopedom(ClassId)
        If ChageNews() Then
            EnoughClassPopedom=defClassPopedomType_Hig
            Exit Function
        End If
        EnoughClassPopedom=-1
        Dim arrPopedomList
            arrPopedomList=Split(AdminClassPopedom,vbCrLf,-1,1)
        Dim arrPopedomItem
        Dim I
        For I=0 To UBound(arrPopedomList)
            arrPopedomItem=Split(arrPopedomList(I),",",-1,1)
            If CLng(arrPopedomItem(0))=CLng(ClassId) Then
                EnoughClassPopedom=CLng(arrPopedomItem(1))
                Exit For
            End If
        Next
    End Function

    '//////////////////////////////////////////////////////////////////
    '//�ж�ָ��pID���ܵĲ���Ȩ���Ƿ����û���Ȩ���б���
    '//����:Flase/True
    Private Function EnoughPopedom(pId)
        EnoughPopedom = ItemInList(pId,Session("AdminPopedom"),",")
    End Function


    '//////////////////////////////////////////////////////////////////
    '//��ǰ�ʻ��ĸ�����Ϣ����    
    Public Property Get AdminLogined()                               '�ʻ��Ƿ��¼
        AdminLogined=Session("AdminLogined")
    End Property
    Public Property Let AdminLogined(byval val)
        Session("AdminLogined")=val
    End Property

    Public Property Get AdminTitle()                                '�ʻ�����
        AdminTitle=Session("AdminTitle")
    End Property
    Public Property Let AdminTitle(byval val)
        Session("AdminTitle")=val
    End Property

    Public Property Get AdminPopedom()                                'Ȩ���б�
        AdminPopedom=Session("AdminPopedom")
    End Property
    Public Property Let AdminPopedom(byval val)
        Session("AdminPopedom")=val
    End Property

    Public Property Get AdminClassPopedom()                         'Ȩ���б�
        AdminClassPopedom=Session("AdminClassPopedom")
    End Property
    Public Property Let AdminClassPopedom(byval val)
        Session("AdminClassPopedom")=val
    End Property

    Public Property Get AdminRoleTitle()                          '������ɫ����
        AdminRoleTitle=Session("AdminRoleTitle")
    End Property
    Public Property Let AdminRoleTitle(byval val)
        Session("AdminRoleTitle")=val
    End Property

    Public Property Get AdminNickName()                           '�ʻ��ǳƣ��������α༭��
        AdminNickName=Session("AdminNickName")
    End Property
    Public Property Let AdminNickName(byval val)
        Session("AdminNickName")=val
    End Property

    Public Property Get AdminTopClassId()                           '����Ա���ɲ鿴�ķ���
        AdminTopClassId=CLng(Session("AdminTopClassId"))
    End Property
    Public Property Let AdminTopClassId(byval val)
        Session("AdminTopClassId")=val
    End Property

    '//////////////////////////////////////////////////////////////////
    '//�������жϵ�ǰ�ʻ�����ӵ�õ�Ȩ������,������Bool��

    '�Ƿ��и��������Ȩ��
    Public Property Get ChagePWD()
        '�����ǰ�ʻ��ǳ�������Ա��ɫ,����������Ȩ��,�����ж�
        If AdminRoleTitle=defAdminRoleTitle Then
            ChagePWD=True
            Exit Property
        End If
        ChagePWD=EnoughPopedom(Podm_ChagePWD)
    End Property

    '�Ƿ����޸�����Ȩ��
    Public Property Get ChageClass()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageClass=True
            Exit Property
        End If
        ChageClass=EnoughPopedom(Podm_ChageClass)
    End Property

    '�Ƿ�������Դ��Դ��Ȩ��
    Public Property Get ChageFrom()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageFrom=True
            Exit Property
        End If
        ChageFrom=EnoughPopedom(Podm_ChageFrom)
    End Property

    '�Ƿ�������Դ���ߵ�Ȩ��
    Public Property Get ChageAuthor()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageAuthor=True
            Exit Property
        End If
        ChageAuthor=EnoughPopedom(Podm_ChageAuthor)
    End Property

    '�Ƿ�����������Դ��Ȩ��
    Public Property Get ChageNews()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageNews=True
            Exit Property
        End If
        ChageNews=EnoughPopedom(Podm_ChageNews)
    End Property

    '�Ƿ�����ջ���վ��Ȩ��
    Public Property Get ChageDustbin()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageDustbin=True
            Exit Property
        End If
        ChageDustbin=EnoughPopedom(Podm_ChageDustbin)
    End Property

   '�Ƿ�����ļ�ϵͳ��Ȩ��
    Public Property Get ManageFiles()
        If AdminRoleTitle=defAdminRoleTitle Then
            ManageFiles=True
            Exit Property
        End If
        ManageFiles=EnoughPopedom(Podm_ManageFiles)
    End Property

    '�Ƿ��й��������ʻ���Ȩ��
    Public Property Get ChangeAdminList()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeAdminList=True
            Exit Property
        End If
        ChangeAdminList=EnoughPopedom(Podm_ChangeAdminList)
    End Property

    '�Ƿ��й����ʻ���ɫ��Ȩ��
    Public Property Get ChangeRole()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeRole=True
            Exit Property
        End If
        ChangeRole=EnoughPopedom(Podm_ChangeRole)
    End Property

    '�Ƿ��и���վ��ҳ���Ȩ��    
    Public Property Get UpdatePage()
        If AdminRoleTitle=defAdminRoleTitle Then
            UpdatePage=True
            Exit Property
        End If
        UpdatePage=EnoughPopedom(Podm_UpdatePage)
    End Property

    '�Ƿ���վ�������滻��Ȩ��
    Public Property Get InsertSYS()
        If AdminRoleTitle=defAdminRoleTitle Then
            InsertSYS=True
            Exit Property
        End If
        InsertSYS=EnoughPopedom(Podm_InsertSYS)
    End Property

    '�Ƿ����������Դ�ļ���Ȩ��
    Public Property Get CreateNewsFile()
        If AdminRoleTitle=defAdminRoleTitle Then
            CreateNewsFile=True
            Exit Property
        End If
        CreateNewsFile=EnoughPopedom(Podm_CreateNewsFile)
    End Property

    '�Ƿ����ƶ���Դ��Ȩ��
    Public Property Get MoveNews()
        If AdminRoleTitle=defAdminRoleTitle Then
            MoveNews=True
            Exit Property
        End If
        MoveNews=EnoughPopedom(Podm_MoveNews)
    End Property

    '�Ƿ�������Դ���Ե�Ȩ��
    Public Property Get ChangeSpeciality()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeSpeciality=True
            Exit Property
        End If
        ChangeSpeciality=EnoughPopedom(Podm_Speciality)
    End Property

    '�Ƿ������۹����Ȩ��
    Public Property Get ChangeCommentList()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeCommentList=True
            Exit Property
        End If
        ChangeCommentList=EnoughPopedom(Podm_CommentList)
    End Property

    '�Ƿ��й�����Դģ���Ȩ��
    Public Property Get ChangeNewsTemplate()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeNewsTemplate=True
            Exit Property
        End If
        ChangeNewsTemplate=EnoughPopedom(Podm_NewsTemplate)
    End Property

    '�Ƿ�������ϵͳ��Ȩ��
    Public Property Get ChangeSysConfig()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeSysConfig=True
            Exit Property
        End If
        ChangeSysConfig=EnoughPopedom(Podm_SysConfig)
    End Property

    '�Ƿ������ݿ��Ȩ��
    Public Property Get ManageDataBase()
        If AdminRoleTitle=defAdminRoleTitle Then
            ManageDataBase=True
            Exit Property
        End If
        ManageDataBase=EnoughPopedom(Podm_ManageDataBase)
    End Property

'////////////////////////////////////////
'  �������������Ȩ������
'////////////////////////////////////////

End Class
%>