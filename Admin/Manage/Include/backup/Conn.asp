<%
'////////////////////////////////////////////////////////////
'//�����������ݿ�����ϵͳ,�ܹ����ͬһ Conn.asp�ļ�Ӧ���ڲ�ͬĿ¼ʱ��������
'//�޷��ҵ����ݿ��ļ�������
'////////////////////////////////////////////////////////////
Option Explicit
'//�����п��ܳ��ֵ��������ݿ�����
Dim DBName(4)
    DBName(1)="../../DataBase/DataBase.mdb"
    DBName(2)="../DataBase/DataBase.mdb"
    DBName(3)="../../DataBase/DataBase.mdb"
    DBName(4)="/Tsys/DataBase/DataBase.mdb"
Dim DBName_Level
	DBName_Level=0
Dim Connstr
'   Connstr="DBQ={dbPath};DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)}" 
    Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath}"
Dim Conn
Set	Conn=Server.CreateObject("ADODB.Connection")
Call ConnOpen()

Function ConnOpen()
    On Error Resume Next
    Dim I
    For I=1 To UBound(DBName)
        Err.Clear
        Conn.Open Replace(Connstr,"{dbPath}",Server.MapPath(DBName(I)))
        If Err.Number=0 Then
            DBName_Level=I
            ConnOpen=true
            Exit Function
        End If
    Next
    If DBName_Level=0 Then
        Err.Clear
        Response.Write "<p>�������ݿ��޷��򿪣���鿴Conn.asp�ļ���*.mdb���ݿ��λ��</p>"
        For I=1 To UBound(DBName)
            '��ӡ�������ݿ�·��
            'Response.Write("" & DBName(I) & "<br>")
        Next
    End If
End Function
%>