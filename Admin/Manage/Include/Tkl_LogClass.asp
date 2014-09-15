<%
'////////////////////////////////////////////////////////////
'//Log系统日制类(TklMilk Boy9732@msn.com)
'//
'//调用实例：
'// Dim TLogClass
'// Set TLogClass=New Tkl_LogClass
'//     TLogClass.AddLog("日制内容")
'// Set TClass=Nothing
'////////////////////////////////////////////////////////////

Class Tkl_LogClass
    Private LogFilePath                         '日制文件路径
    Private LogFileMaxSize                      '日制文件最大限制(byte)
    Private Fso                                 'Fso对象
    Private Fle                                 '日制文件对象
    Private Sub Class_Initialize
        LogFilePath=Server.MapPath("./TsysLog.txt")
        LogFileMaxSize=1024*1024*10             '10M
        Set Fso=Server.CreateObject("Scripting.FileSystemObject")
        If Fso.FileExists(LogFilePath) Then
            Set Fle=Fso.GetFile(LogFilePath)
            '日制文件超出限制大小则清空
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