<%
Sub UpdateAdminTime()
    Dim Coll
    Set Coll = New UserInfo_Collection_Class
        Coll.Refh()
    Dim myInfo
    Set myInfo = Coll.GetUser(Session("AdminTitle"))
    If myInfo.Name <> "" Then
        '更新当前管理员的最后登录时间
        myInfo.UpdateTime()
        Coll.Add(myInfo)
    End If
End Sub
%>