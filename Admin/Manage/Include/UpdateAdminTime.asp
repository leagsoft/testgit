<%
Sub UpdateAdminTime()
    Dim Coll
    Set Coll = New UserInfo_Collection_Class
        Coll.Refh()
    Dim myInfo
    Set myInfo = Coll.GetUser(Session("AdminTitle"))
    If myInfo.Name <> "" Then
        '���µ�ǰ����Ա������¼ʱ��
        myInfo.UpdateTime()
        Coll.Add(myInfo)
    End If
End Sub
%>