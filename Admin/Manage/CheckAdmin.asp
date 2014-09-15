<%
'///////////////////////////////////////////
'//判断登录管理员信息
'//参数：登录名,密码(已由MD5加密)
'//返回：预定义数据格式：
'//     "名称,权限列表,角色名称,管理员昵称,分类权限列表,限制仅可查看的分类"
'//     注：当管理员被锁定则返回：{LOCK}
Function CheckAdmin(mTitle,mPwd)
    Dim mRs
    'Sql Server语句中好像不能用UCase(函数)
    '**Set mRs=Conn.ExeCute("Select Title,Popedom,RoleTitle,Lock,NickName,ClassPopedom,ClassId From View_AdminInfo Where UCase(Title)='" & UCase(mTitle) & "' And Pwd='" & mPwd &"'")
    Set mRs=Conn.ExeCute("Select Title,Popedom,RoleTitle,Lock,NickName,ClassPopedom,ClassId From View_AdminInfo Where Title='" & UCase(mTitle) & "' And Pwd='" & mPwd &"'")
    Dim mResult
    If mRs.Eof And mRs.Bof Then
        mResult=""
    Else
        If CBool(mRs("Lock")) Then
            mResult="{LOCK}"
        Else
            mResult=mRs("Title") & vbTab & mRs("Popedom") & vbTab & mRs("RoleTitle") & vbTab & mRs("NickName") & vbTab & mRs("ClassPopedom") & vbTab & mRs("ClassId")
        End If
    End If
    mRs.Close
    Set mRs=Nothing
    CheckAdmin = mResult
End Function
%>