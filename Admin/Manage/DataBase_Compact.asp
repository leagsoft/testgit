<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not CBool(SysAdmin.Logined) Then
    Response.Redirect("Login.asp")
End If

%>
<html>
<head>
<title>DataBase_Compact</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link href="Include/ManageStyle.css" rel="stylesheet" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>

<body>
<%
Dim Work
    Work=Request("Work")
Select Case Work
    Case "DoCompact"
        Call DoCompact()
    Case "BakDB"
        Call BakDB()
    Case "DoBakDB"
        Call DoBakDB()
    Case "ReBakDB"
        Call ReBakDB()
    Case "ExeCuteSql"
        Call ExeCuteSql()
    Case "DoExeCuteSql"
        Call DoExeCuteSql()
    Case Else
        Call CompactDB()
End Select
%>
<%Sub CompactDB()%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td align="center" class="BarTitleBg">ѹ�� ���ݿ�</td>
  </tr>
  <tr>
    <td height="105" align="center" class="BarTitle"><form name="form1" method="post" action="?">
        <table width="75%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="32" align="center" class="BarText">ע�����ڶ�Access���ݽ���ѹ������,���Լ���ؼ�С���ļ���С,������ݶ�ȡ����...<br>
              ������ѹ�����ݿ�ǰ�����ݿ����һ��<a href="?Work=BakDB">����</a>��</td>
          </tr>
          <tr>
            <td align="center"><input name="Submit" type="submit" class="button01-out" value="ȷ  ��">
              <input name="Submit3" type="button" class="button01-out" value="��  ��" onclick="window.history.back();">
              <input name="Work" type="hidden" id="Work4" value="DoCompact"> </td>
          </tr>
        </table>
        </form></td>
  </tr>
</table>
<%End Sub%>
<%Sub BakDB()%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td align="center" bgcolor="#CCCCCC" class="BarTitleBg">���� ���ݿ�</td>
  </tr>
  <tr> 
    <td height="105" align="center" bgcolor="#F6f6f6"><form action="?" method="post" name="form2" id="form2">
        <table width="75%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="27" align="center" class="BarText">ע��ϵͳ�Զ�����Դ���ݿⱸ��Ϊԭ���ݿ�ͬĿ¼����Ϊ <font color="#006600">DataBase.mdb.bak</font> 
              ���ļ�,ͬʱ�Ḵ���ϴεı����ļ�</td>
          </tr>
          <tr>
            <td align="center"><input name="Submit2" type="submit" class="button01-out" value="ȷ  ��"> 
              <input name="Submit32" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"> 
              <input name="Work" type="hidden" id="Work22" value="DoBakDB"> 
            </td>
          </tr>
        </table>
        
      </form></td>
  </tr>
</table>
<%End Sub%>
<%Sub ReBakDB()%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td align="center" bgcolor="#CCCCCC" class="BarTitleBg">��ԭ ���ݿ�</td>
  </tr>
  <tr> 
    <td height="105" align="center" bgcolor="#F6f6f6"><table width="75%" height="80" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="35" align="center" class="BarText">Ϊȷ������������,ϵͳ�ݲ��ṩ�˲����������ֶ����л�ԭ<br>
            �����������ݿ�ͬ��Ŀ¼�е�.bak���ݼ��ļ�����Ϊ.mdb�ȿ�</td>
        </tr>
        <tr>
          <td align="center"> 
            <input name="Submit33" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"></td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
<%End Sub%>
<%Sub ExeCuteSql()%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="ContentTabBg">
  <tr> 
    <td align="center" bgcolor="#CCCCCC" class="BarTitleBg">ִ��Sql�ű�</td>
  </tr>
  <tr> 
    <td height="105" align="center" bgcolor="#F6f6f6" class="BarTitle"><form action="?" method="post" name="form4" id="form4">
        <table width="90%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="16" class="BarText">ע��һ��ֻ��ִ��һ��Sql���</td>
          </tr>
          <tr>
            <td height="300" align="center" valign="top"> 
                <textarea name="Content" rows="5" wrap="OFF" class="Input" id="Content" style="width:100%;height:100%"></textarea>          
              </td>
          </tr>
          <tr> 
            <td align="center"><input name="Submit4" type="submit" class="button01-out" value="ȷ  ��"> 
              <input name="Submit22" type="reset" class="button01-out" value="��  ԭ"> 
              <input name="Submit34" type="button" class="button01-out" value="��  ��" onclick="window.history.back();"> 
              <input name="Work" type="hidden" id="Work" value="DoExeCuteSql"> 
            </td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<%End Sub%>
</body>
</html>
<%
'////////////////////////////
'//ѹ�����ݿ�
Sub DoCompact()
    If Not SysAdmin.ManageDataBase Then
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim I
    Dim TargetDB,ResourceDB
    Dim oJetEngine
    Dim Fso
    Const Jet_Conn_Partial = "Provider=Microsoft.Jet.OLEDB.4.0; Data source="
    Set oJetEngine = Server.CreateObject("JRO.JetEngine")
    Set Fso= Server.CreateObject(FsoObjectStr)
    '�ر����ݿ�����
    Conn.Close
    Set Conn=Nothing

    For I=1 To UBound(DBName)
        ResourceDB=Server.MapPath(DBName(I))

        If Fso.FileExists(ResourceDB) Then
            '������ʱ�ļ�
            TargetDB=Server.MapPath(DBName(I)&".bak")
            If Fso.FileExists(TargetDB) Then
                Fso.DeleteFile(TargetDB)
            End If
            oJetEngine.CompactDatabase Jet_Conn_Partial&ResourceDB,Jet_Conn_Partial&TargetDB
            Fso.DeleteFile ResourceDB
            Fso.MoveFile TargetDB,ResourceDB
        End If
    Next

    Set Fso=Nothing
    Set oJetEngine=Nothing

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "ѹ�����ݿ�")
    Set LogClass=Nothing

    Response.Write("<script>alert(""<�����ɹ�>\n���ݿ�ѹ�����"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End Sub

'////////////////////////////
'//�������ݿ�
Sub DoBakDB()
    If Not SysAdmin.ManageDataBase Then
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim I
    Dim TargetDB,ResourceDB
    Dim Fso
    Set Fso= Server.CreateObject(FsoObjectStr)
    '�ر����ݿ�����
    Conn.Close
    Set Conn=Nothing

    For I=1 To UBound(DBName)
        ResourceDB=Server.MapPath(DBName(I))
        If Fso.FileExists(ResourceDB) Then
            'ԭ���ݿ����Ƿ����
            If Fso.FileExists(ResourceDB) Then
                '����ԭ�����ļ�
                TargetDB=Server.MapPath(DBName(I)&".bak")
                Fso.CopyFile ResourceDB,TargetDB,True
            End If
        End If
    Next

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "�������ݿ�")
    Set LogClass=Nothing

    Set Fso=Nothing
    Response.Write("<script>alert(""<�����ɹ�>\n���ݿⱸ�����"& SoftCopyright_Script &""");window.history.back();</script>")
    Response.End()
End Sub

'////////////////////////////
'//ִ��Sql����
Sub DoExeCuteSql()
    If Not SysAdmin.ManageDataBase Then
        Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
        Response.End()
    End If

    Dim Content
        Content=Request("Content")
    Err.Clear
    On Error Resume Next
    Conn.BeginTrans
    Conn.ExeCute(Content)
    If Err.Number<>0 Then
        Conn.Rollback
        Response.Write("<script>alert(""<����ʧ��>\nSql��������뷵�ز鿴"& SoftCopyright_Script &""");window.history.back();</script>")

    Else
        Conn.CommitTrans
        Response.Write("<script>alert(""<�����ɹ�>\nSql���ִ�����"& SoftCopyright_Script &""");window.history.back();</script>")
    End If

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "ִ��Sql���")
    Set LogClass=Nothing

    Response.End
End Sub
%>