<%
'///////////////////////////////////////////
'//������ȡ�õ�ǰId��������·��
'//��������ǰ�ڵ�Id,����?�������Ĳ���ֵ
'//���أ��ַ���������������� > �������� > ������� > ����������
Function GetClassPath(Id,Url)
    Dim Sql,str
    Dim tId,Counter
        tId=Id
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
    do
        Sql="Select Id,Title,Parent From ClassList Where Id="&tId
        Set Rs=Conn.ExeCute(Sql)
        If Not(Rs.Eof And Rs.Bof)Then
            str=" > <a href='"&Url&"Parent="&Rs("Id")&"'>"&Rs("Title")&"</a>"&str
            tId=CInt(Rs("Parent"))
        Else
            StopRun=True
        End If
        Rs.Close
    Loop Until(StopRun Or tId=0)
    str="<a href='"&Url&"Parent=0'>�����</a>" & str
    GetClassPath=str
End Function

'///////////////////////////////////////////
'//������ȡ�õ�ǰId��������·��
'//��������ǰ�ڵ�Id,����?�������Ĳ���ֵ
'//���أ��ַ���������������� > �������� > ������� > ����������
Function GetClassPath2(RootId,Id,Url)
    Dim Rs,Str
    Set Rs=Conn.ExeCute("Select Id,Parent,Title From ClassList Where Id="&Id&" And Id<>"&RootId)
	If Not(Rs.Eof And Rs.Bof) Then
		If Url="" Then
			Str= " &gt; <a href=""?Parent="&Rs("Id")&""">" & Rs("Title") & "</a>"
		Else
			Str= " &gt; <a href="""&Url&"Parent="&Rs("Id")&""">" & Rs("Title") & "</a>"			
		End If
		Str=GetClassPath2(RootId,Rs("Parent"),Url) & Str
	Else
		If Url="" Then
			Str="<a href=""?Parent="&RootId&""">�����</a>" & Str			
		Else
			Str="<a href="""&Url&"Parent="&RootId&""">�����</a>" & Str
		End If
	End If
	Rs.Close
	Set Rs=Nothing
	GetClassPath2=Str
End Function

'///////////////////////////////////////////
'//�������ݹ��ѵ�ǰĿ¼�����е��¼�Ŀ¼�����޲���Ŀ¼��
'//��������ĿId
Function AllChildClass(Parent)
    Dim Rs
    Set Rs=Conn.ExeCute("Select Id From ClassList Where Parent="&Parent)
    While Not Rs.Eof
        AllChildClass=AllChildClass & "," & Rs("Id")
        AllChildClass=AllChildClass & AllChildClass(Rs("Id"))
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Function

'//////////////////////////////////////////////
'//�����������Դָ�����ݿ��ֶε���Ϣ
'//�������ֶ���,��ԴId
'//���أ����ݿ��ֶ�ֵ
Function GetNewsFieldValue(FieldName,Id)
    Dim Rs,Sql
        Sql="Select "&FieldName&" From view_AllNewsInfo Where Id="&Id
    Set Rs=Conn.ExeCute(Sql)
    If Not(Rs.Eof And Rs.Bof) Then
        GetNewsFieldValue=Rs(FieldName)
    Else
        GetNewsFieldValue=Null
    End If
    Rs.Close
    Set Rs=Nothing
End Function
%>