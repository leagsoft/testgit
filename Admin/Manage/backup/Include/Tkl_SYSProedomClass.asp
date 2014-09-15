<%
'////////////////////////////////////////////////////////////
'//资源系统管理员信息类(TklMilk Boy9732@msn.com)
'//用途：对管理员的信息读取,权限判断进行规范，统一操作
'//调用实例：
'//Dim SysAdmin
'//Set SysAdmin=New SYSProedom_Class
'//Set SysAdmin=Nothing
'//属性：
'//------------属性名-------------返回类型------------------简介-----
'//         .Logined             Bool                管理员是否登录
'//         .AdminTitle          String              管理员帐户名称
'//         .AdminNickName       String              管理员昵称
'//         .AdminRoleTitle      String              管理员所属角色名称
'//         更多请详见其它代码注解...
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
'  在这里添加您的权限变量定义
'////////////////////////////////////////

    Private Sub Class_Initialize
        '初始化权限项目唯一标识
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
'  在这里初始化您的权限常量
'////////////////////////////////////////

        '初始化权限列表
        PopedomList=Podm_ChagePWD&",修改密码,"&_
                    Podm_ChageClass&",修改栏目,"&_
                    Podm_ChageFrom&",修改来源,"&_
                    Podm_ChageAuthor&",修改作者,"&_
                    Podm_ChageNews&",管理所有资源,"&_
                    Podm_ChageDustbin&",回收站管理,"&_
                    Podm_ChangeCommentList&",评论管理,"&_
                    Podm_ManageFiles&",文件系统管理,"&_
                    Podm_UpdatePage&",站点页面更新,"&_
                    Podm_InsertSYS&",站点内容替换,"&_
                    Podm_CreateNewsFile&",生成资源,"&_
                    Podm_MoveNews&",移动资源,"&_
                    Podm_Speciality&",资源特性,"&_
                    Podm_CommentList&",评论管理,"&_
                    Podm_NewsTemplate&",资源模版,"&_
                    Podm_ChangeAdminList&",用户管理,"&_
                    Podm_ChangeRole&",角色操作,"&_
                    Podm_SysConfig&",系统设置,"&_
                    Podm_ManageDataBase&",数据库管理,"&_
                    ""

    End Sub

    Public Property Get defPopedomList()
        defPopedomList=PopedomList
    End Property

    '默认的系统超级管理员的名称,该用户不可被删除和修改
    Public Property Get defAdminUserTitle()
        defAdminUserTitle="Admin"
    End Property

    '名称AdminRoleTitle的角色不可被删除，和修改
    Public Property Get defAdminRoleTitle()
        defAdminRoleTitle="超级管理员"
    End Property

    '栏目权限类型,低
    Public Property Get defClassPopedomType_Low()
        defClassPopedomType_Low=0
    End Property
    '栏目权限类型,中
    Public Property Get defClassPopedomType_Mid()
        defClassPopedomType_Mid=1
    End Property
    '栏目权限类型,高
    Public Property Get defClassPopedomType_Hig()
        defClassPopedomType_Hig=2
    End Property

    '管理员是否登录,返回bool
    Public Property Get Logined()
        If Session("AdminLogined")="TRUE" And Trim(Session("AdminTitle"))<>"" Then
            Logined=true
        Else
            Logined=false
        End If
    End Property
    '管理员退出
    Public Sub LogOut()
        Session.Abandon
    End Sub

    '////////////////////////////////////////////////////////
    '//判断mItem是否在stritemList列表中，此列表由splitStr为分隔符
    '例：3 是否 1,2,3,4,5
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
    '//获得指定分类的的操作权值,
    '//返回：若有则返回权值;无权限则返回-1
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
    '//判断指定pID功能的操作权限是否在用户的权限列表当中
    '//返回:Flase/True
    Private Function EnoughPopedom(pId)
        EnoughPopedom = ItemInList(pId,Session("AdminPopedom"),",")
    End Function


    '//////////////////////////////////////////////////////////////////
    '//当前帐户的各项信息属性    
    Public Property Get AdminLogined()                               '帐户是否登录
        AdminLogined=Session("AdminLogined")
    End Property
    Public Property Let AdminLogined(byval val)
        Session("AdminLogined")=val
    End Property

    Public Property Get AdminTitle()                                '帐户名称
        AdminTitle=Session("AdminTitle")
    End Property
    Public Property Let AdminTitle(byval val)
        Session("AdminTitle")=val
    End Property

    Public Property Get AdminPopedom()                                '权限列表
        AdminPopedom=Session("AdminPopedom")
    End Property
    Public Property Let AdminPopedom(byval val)
        Session("AdminPopedom")=val
    End Property

    Public Property Get AdminClassPopedom()                         '权限列表
        AdminClassPopedom=Session("AdminClassPopedom")
    End Property
    Public Property Let AdminClassPopedom(byval val)
        Session("AdminClassPopedom")=val
    End Property

    Public Property Get AdminRoleTitle()                          '所属角色名称
        AdminRoleTitle=Session("AdminRoleTitle")
    End Property
    Public Property Let AdminRoleTitle(byval val)
        Session("AdminRoleTitle")=val
    End Property

    Public Property Get AdminNickName()                           '帐户昵称（用于再任编辑）
        AdminNickName=Session("AdminNickName")
    End Property
    Public Property Let AdminNickName(byval val)
        Session("AdminNickName")=val
    End Property

    Public Property Get AdminTopClassId()                           '管理员仅可查看的分类
        AdminTopClassId=CLng(Session("AdminTopClassId"))
    End Property
    Public Property Let AdminTopClassId(byval val)
        Session("AdminTopClassId")=val
    End Property

    '//////////////////////////////////////////////////////////////////
    '//以下是判断当前帐户的所拥用的权限属性,均返回Bool型

    '是否有更改密码的权限
    Public Property Get ChagePWD()
        '如果当前帐户是超级管理员角色,则所有所有权限,无需判断
        If AdminRoleTitle=defAdminRoleTitle Then
            ChagePWD=True
            Exit Property
        End If
        ChagePWD=EnoughPopedom(Podm_ChagePWD)
    End Property

    '是否有修改类别的权限
    Public Property Get ChageClass()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageClass=True
            Exit Property
        End If
        ChageClass=EnoughPopedom(Podm_ChageClass)
    End Property

    '是否有修资源来源的权限
    Public Property Get ChageFrom()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageFrom=True
            Exit Property
        End If
        ChageFrom=EnoughPopedom(Podm_ChageFrom)
    End Property

    '是否有修资源作者的权限
    Public Property Get ChageAuthor()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageAuthor=True
            Exit Property
        End If
        ChageAuthor=EnoughPopedom(Podm_ChageAuthor)
    End Property

    '是否有修所有资源的权限
    Public Property Get ChageNews()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageNews=True
            Exit Property
        End If
        ChageNews=EnoughPopedom(Podm_ChageNews)
    End Property

    '是否有清空回收站的权限
    Public Property Get ChageDustbin()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChageDustbin=True
            Exit Property
        End If
        ChageDustbin=EnoughPopedom(Podm_ChageDustbin)
    End Property

   '是否管理文件系统的权限
    Public Property Get ManageFiles()
        If AdminRoleTitle=defAdminRoleTitle Then
            ManageFiles=True
            Exit Property
        End If
        ManageFiles=EnoughPopedom(Podm_ManageFiles)
    End Property

    '是否有管理所有帐户的权限
    Public Property Get ChangeAdminList()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeAdminList=True
            Exit Property
        End If
        ChangeAdminList=EnoughPopedom(Podm_ChangeAdminList)
    End Property

    '是否有管理帐户角色的权限
    Public Property Get ChangeRole()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeRole=True
            Exit Property
        End If
        ChangeRole=EnoughPopedom(Podm_ChangeRole)
    End Property

    '是否有更新站点页面的权限    
    Public Property Get UpdatePage()
        If AdminRoleTitle=defAdminRoleTitle Then
            UpdatePage=True
            Exit Property
        End If
        UpdatePage=EnoughPopedom(Podm_UpdatePage)
    End Property

    '是否有站点内容替换的权限
    Public Property Get InsertSYS()
        If AdminRoleTitle=defAdminRoleTitle Then
            InsertSYS=True
            Exit Property
        End If
        InsertSYS=EnoughPopedom(Podm_InsertSYS)
    End Property

    '是否具有生成资源文件的权限
    Public Property Get CreateNewsFile()
        If AdminRoleTitle=defAdminRoleTitle Then
            CreateNewsFile=True
            Exit Property
        End If
        CreateNewsFile=EnoughPopedom(Podm_CreateNewsFile)
    End Property

    '是否有移动资源的权限
    Public Property Get MoveNews()
        If AdminRoleTitle=defAdminRoleTitle Then
            MoveNews=True
            Exit Property
        End If
        MoveNews=EnoughPopedom(Podm_MoveNews)
    End Property

    '是否有修资源特性的权限
    Public Property Get ChangeSpeciality()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeSpeciality=True
            Exit Property
        End If
        ChangeSpeciality=EnoughPopedom(Podm_Speciality)
    End Property

    '是否有评论管理的权限
    Public Property Get ChangeCommentList()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeCommentList=True
            Exit Property
        End If
        ChangeCommentList=EnoughPopedom(Podm_CommentList)
    End Property

    '是否有管理资源模板的权限
    Public Property Get ChangeNewsTemplate()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeNewsTemplate=True
            Exit Property
        End If
        ChangeNewsTemplate=EnoughPopedom(Podm_NewsTemplate)
    End Property

    '是否有配制系统的权限
    Public Property Get ChangeSysConfig()
        If AdminRoleTitle=defAdminRoleTitle Then
            ChangeSysConfig=True
            Exit Property
        End If
        ChangeSysConfig=EnoughPopedom(Podm_SysConfig)
    End Property

    '是否有数据库的权限
    Public Property Get ManageDataBase()
        If AdminRoleTitle=defAdminRoleTitle Then
            ManageDataBase=True
            Exit Property
        End If
        ManageDataBase=EnoughPopedom(Podm_ManageDataBase)
    End Property

'////////////////////////////////////////
'  在这里添加您的权限属性
'////////////////////////////////////////

End Class
%>