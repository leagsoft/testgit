<%
'///////////////////////////////////////////
'//ȡ�õ�ǰId��������·��
'//��������ǰ�ڵ�Id,����?�������Ĳ���ֵ
'//���أ��ַ���������������� > ����1 > ����1-1 > ����1-1-1��
Function Spec_GetSpecialityPath(Id,Url)
    Dim Sql,str
    Dim tId
        tId=Id
    Dim Counter
        Counter=0
    Dim StopRun
        StopRun=false
    If Instr(Url,"?")=0 Then
        Url=Url&"?"
    Else
        If Right(Url,1)<>"?" Then
            Url=Url&"&"
        End If
    End If
    Dim Rs
    Do
        Sql="Select Id,Title,Parent From News_Speciality Where Id="&tId
        Set Rs=Conn.ExeCute(Sql)
        If Not(Rs.Eof And Rs.Bof)Then
            str=" > <a href='"&Url&"Parent="&Rs("Id")&"'>"&Rs("Title")&"</a>"&str
            tId=CInt(Rs("Parent"))
        Else
            StopRun=true
        End If
        Rs.Close
    Loop Until(StopRun Or tId=0)
    str="<a href='"&Url&"Parent=0'>�����</a>" & str
    Spec_GetSpecialityPath=str
End Function

'///////////////////////////////////////////
'//�ݹ��ѵ�ǰĿ¼�����е��¼�Ŀ¼�����޲���Ŀ¼��
Function Spec_AllChildClass(Parent)
    Dim Sql
        Sql="Select Id From News_Sepciality Where Parent="&Parent
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        Spec_AllChildClass=Spec_AllChildClass&","&Rs("Id")
        Spec_AllChildClass=Spec_AllChildClass&AllChildClass(Rs("Id"))
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Function
%>