<%
'///////////////////////////////////////////
'//�жϵ�¼����Ա��Ϣ
'//��������¼��,����(����MD5����)
'//���أ�Ԥ�������ݸ�ʽ��
'//     "����,Ȩ���б�,��ɫ����,����Ա�ǳ�,����Ȩ���б�,���ƽ��ɲ鿴�ķ���"
'//     ע��������Ա�������򷵻أ�{LOCK}
Function CheckAdmin(mTitle,mPwd)
    Dim mRs
    'Sql Server����к�������UCase(����)
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