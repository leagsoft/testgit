<%
'///////////////////////////////////////////
'//取得当前Id所在树的路径
'//参数：当前节点Id,链接?后所带的参数值
'//返回：字符串，（例：根类别 > 特性1 > 特性1-1 > 特性1-1-1）
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
    str="<a href='"&Url&"Parent=0'>根类别</a>" & str
    Spec_GetSpecialityPath=str
End Function

'///////////////////////////////////////////
'//递归搜当前目录下所有的下级目录（无限层型目录）
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