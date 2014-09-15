<%
'////////////////////////////////////////////////////////////
'//　　自能数据库链接系统,能够解决同一 Conn.asp文件应用于不同目录时所产生的
'//无法找到数据库文件的问题
'////////////////////////////////////////////////////////////
Option Explicit
'//配制有可能出现的所有数据库链接
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
        Response.Write "<p>以下数据库无法打开，请查看Conn.asp文件及*.mdb数据库的位置</p>"
        For I=1 To UBound(DBName)
            '打印所有数据库路径
            'Response.Write("" & DBName(I) & "<br>")
        Next
    End If
End Function
%>