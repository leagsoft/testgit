<%
'////////////////////////////////////////////////////////////
'//Logϵͳ������(TklMilk Boy9732@msn.com)
'//
'//����ʵ����
'// Dim TLogClass
'// Set TLogClass=New Tkl_LogClass
'//     TLogClass.AddLog("��������")
'// Set TClass=Nothing
'////////////////////////////////////////////////////////////

Class Tkl_LogClass
    Private LogFilePath                         '�����ļ�·��
    Private LogFileMaxSize                      '�����ļ��������(byte)
    Private Fso                                 'Fso����
    Private Fle                                 '�����ļ�����
    Private Sub Class_Initialize
        LogFilePath=Server.MapPath("./TsysLog.txt")
        LogFileMaxSize=1024*1024*10             '10M
        Set Fso=Server.CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(LogFilePath) Then
            Set Fle=Fso.GetFile(LogFilePath)
            '�����ļ��������ƴ�С�����
            If Fle.Size>=LogFileMaxSize Then
                Fso.DeleteFile(LogFilePath)
            End If
        End If
    End Sub

    Public Function AddLog(str)
'	response.write "<font color=white>" & LogFilePath & str& "</font>"
'response.end
 '       Set Fle=Fso.OpenTextFile(LogFilePath,8,True)
  '      Fle.Writeline(Now() & " " & str)
   '     Fle.Close
    End Function

    Private Sub Class_Terminate
        Set Fso=Nothing
    End Sub
End Class
%>