<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<!--#include file="Include/ClassList_Fun.asp" -->
<!--#include file="Include/CreateFile_Fun.asp" -->
<!--#include file="Include/Tkl_StringClass.asp" -->
<!--#Include File="Include/Tkl_LogClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
'If Not CBool(SysAdmin.Logined) Then
    'Response.Redirect("Login.asp")
'End If
%>
<html>
<head>
<title>News_Mdy.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<script src="Include/Tkl_Skin.js"></script>
</head>

<body>
<%
Select Case Request("Work")
    Case "ClearDustbin"
        Call ClearDustbin()
    Case "SaveMdy"
        Call SaveMdy()
    Case "DelReco"
        Call DelReco()
    Case "AddReco"
        Call AddReco()
    Case "CheckReco"
        Call CheckReco()
    Case "CreateSelectedFile"
        Call CreateSelectedFile()
    Case Else
End Select
%>
</body>
</html>
<%
'///////////////////////////
'//�����޸ĵ���Դ��¼
Sub SaveMdy()

    Dim Id,ClassId,Title,Url,Author,From,KeyWord
    Dim Editor,Count,Speciality
    Dim Content,ImgNews,SmallImg,BigImg,ShortContent,NowTime
    Dim CreateFile
    'ȡ�ò���
    Id=CLng(Request("Id"))
    ClassId=Request("radioBoxItem")
    If ClassId="" Then
        Response.Write("<script>alert(""������[��Դ���]"");window.history.back();</script>")
        Response.End
    End If
    '**** Add By BennyLIu:20040625   '���������
    '**If Session("QXMC")="����ͳ����Ϣ" then
	If Session("QXMC")="����ͳ����Ϣ" or Session("QXMC")="�־ֶ�̬" then		'Modify By BennyLiu:20040712
		Browser=Request("Browser")					'ȡ�������				
		DocumentType=Trim(Request("DocumentType")) 'ȡ���ļ�����
	End If
	'*** End Add *********
    Title=Request("Title")
    Url=Request("Url")
    Author=Request("Author")
    if Author="" then
		Author="1"
	end if
    From="3"
    
    KeyWord=Request("KeyWord")
    Editor=Request("Editor")
    FromBM=Request("From")
    Count=Request("Count")
    Content=Request("NewsContent")
    ImgNews=CBool(Request("ImgNews"))
    SmallImg = Request("SmallImg")
    NowTime=Now
    '//���
    Dim Sql
    Sql="Select Top 1 * From News Where Id=" & ID & " Order By ID DESC"
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")    
    Rs.Open Sql,Conn,1,3
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("<script>alert(""��¼������"");window.history.back();</script>")
        Response.End
    End If
    Rs("Class")=ClassId
    Rs("Title")=Title
    Rs("Url")=Url
    'Modify By BennyLIU:20040625
    
    if Session("QXMC")="�ֳ�ר��" then  
		Rs("Author")=Author
    end if 
    'End Modify

    Rs("KeyWord")=KeyWord
    Rs("Editor")=Editor
    Rs("FromBM")=FromBM'Modify By Housetea:20050304
    Rs("Count")=Count
    Rs("Content")=Content
    If ImgNews Then
        Rs("ImgNews")=1
    Else
        Rs("ImgNews")=0
    End If
    Rs("SmallImg")=SmallImg
	If Def_ReCheckAfterModify Then
		Rs("IsChecked")=0
	Else
		Rs("IsChecked")=1
	End If
    Rs("Created")=0
    Rs("UpTime")=NowTime
    'Add By BennyLiu:20040625
	if Browser<>"" then
		Rs("Browser")=Browser
	end if
	if DocumentType<>"" then
		Rs("IsDocument")=DocumentType
	end if
	'End Add
    Rs.Update
    Rs.Close
    Set Rs=Nothing

    '���
    Response.Redirect("News_List.asp?Work="&Request.Cookies("ZGW_NewsSys")("News_List_WorkType"))
End Sub


'///////////////////////////
'//�����Դ��¼
Sub AddReco()
    Dim ClassId,Title,Url,Author,From,KeyWord
    Dim Editor,Count
    Dim Content,ImgNews,SmallImg,BigImg,ShortContent,NowTime,UserBM
    'ȡ�ò���
    ClassId=Request("radioBoxItem")
    If ClassId="" Then
        Response.Write("<script>alert(""������[��Դ���]"");window.history.back();</script>")
        Response.End
    End If
    'Add By BennyLiu:20040628��Ϊ�˶��������Ȩ�޶�����
	'**If session("QXMC")="����ͳ����Ϣ" then
	If session("QXMC")="����ͳ����Ϣ" or Session("QXMC")="�־ֶ�̬" then		'Modify By BennyLiu:20040712
		Browser=Request("Browser")		'ȡ�ÿ������
		DocumentType=Trim(Request("DocumentType"))	'ȡ���ļ�����
	End If
	'End Add
    Title=Request("Title")
    Url=Request("Url")
    Author=Request("Author")		'��Դ���ߣ�����ֳ������
    From="4"
    FromBM=request("From")
    KeyWord=Request("KeyWord")
    Editor=Request("Editor")
    Count=Request("Count")
    Content=Request("NewsContent")
    ImgNews=CBool(Request("ImgNews"))
    SmallImg = Request("SmallImg")
    UserBM=	Request("From")
    NowTime=Now

    Dim Sql
        Sql="Select Top 1 * From News Order By ID DESC"
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
    Rs.Open Sql,Conn,1,3
    Rs.AddNew
    '���
    Rs("Class")=ClassId
    Rs("Title")=Title
    Rs("Url")=Url
    'Modify By BennyLiu:20040625
    If Session("QXMC")="�ֳ�ר��" then
		Rs("Author")=Author
    End If
    'End Modify
    Rs("From")=From
    Rs("FromBM")=FromBM
    Rs("KeyWord")=KeyWord
    Rs("Editor")=Editor
    Rs("Count")=Count
    Rs("Content")=Content
    If ImgNews Then
        Rs("ImgNews")=1
    Else
        Rs("ImgNews")=0
    End If
    Rs("SmallImg")=SmallImg
    Rs("AddTime")=NowTime
    Rs("UpTime")=NowTime
	Rs("UserBM")=UserBM
	If Browser<>"" then		'����߽���
		Rs("Browser")=Browser
	End If
	If DocumentType<>"" then
		Rs("IsDocument")=DocumentType		'�ļ����ͽ���
	End If
    Rs.Update
    Rs.Close
    Set Rs=Nothing
    '����һ��Session�������´�¼����Ϣʱ������������ͬ�������
    If Browser<>"" then
		Session("Browser")=Browser
		Session.Timeout =30
    End If
    Response.Write("<script>if(confirm(""<��ӳɹ�>\n�Ƿ�����������Դ��"")){window.location='News_Add.asp?Work=AddReco'}else{window.location='"&"News_List.asp?Work="&Request.Cookies("ZGW_NewsSys")("News_List_WorkType")&"'}</script>")
	If CBool(Request("CurrentClassIdUsed")) Then
		Response.Cookies("ZGW_NewsSys")("CurrentClassIdUsed")=ClassId
	Else
		Response.Cookies("ZGW_NewsSys")("CurrentClassIdUsed")=""
	End If
End Sub

'//////////////////
'//��ʵ,����ɾ��
Sub DelReco()
    '��ǰ����Ա����Ȩ�޵ȼ�ֵ
    Dim ClassPopedomType
    Dim Id,IdList,FinishedNum
        FinishedNum=0
        Id=Request("Id")
        IdList=Split(Id,",",-1,1)
    Dim I,RealDel
        RealDel=CBool(Request("RealDel"))
    Dim Sql,Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")
    For I=0 To UBound(IdList)
        ClassPopedomType=SysAdmin.EnoughClassPopedom(GetNewsFieldValue("Class",CLng(IdList(I))))
        '�Ƿ�߱��㹻��Ȩ��
        'If ClassPopedomType=SysAdmin.defClassPopedomType_Hig Then
            Sql="Select * From News Where Id=" & CLng(IdList(I))
            Rs.Open Sql,Conn,1,3
            If Not(Rs.Eof And Rs.Bof) Then
                'ɾ�������ļ�
                DeleteNewsFile Rs("FilePath"),Rs("Id")
                If RealDel Then
                    Rs.Delete
                    Rs.Update
                Else
                    Rs("Created")=0
                    If CBool(Rs("Del")) Then
                        Rs("Del")=0
                    Else
                        Rs("Del")=1
                    End If
                    Rs.Update
                End If
                Rs.Close
                FinishedNum=FinishedNum+1
            End If
       'End If
    Next
    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "ɾ��"&UBound(IdList)+1&"����Դ,ʵ�ʳɹ�"&FinishedNum&"��")
    Set LogClass=Nothing
    Response.Write("<script>alert(""<ϵͳ��ʾ>\n�˴β���Ԥ�Ʋ�����"&UBound(IdList)+1&"��������,ʵ�ʳɹ�������"&FinishedNum&"������δ��ɣ�"&(UBound(IdList)+1-FinishedNum)&"����"");window.location="""&"News_List.asp?Work="&Request.Cookies("ZGW_NewsSys")("News_List_WorkType")&"&CurrentPage="&Request.Cookies("ZGW_NewsSys")("News_List_CurrentPage")&""";</script>")
    Response.End
End Sub


'/////////////////
'//��ջ���վ
Sub ClearDustbin()
    
    Dim Sql
        Sql="DELETE FROM News WHERE DEL=1"
    Conn.ExeCute(Sql)

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "��ջ���վ")
    Set LogClass=Nothing

    Response.Redirect "News_List.asp?Work="&Request.Cookies("ZGW_NewsSys")("News_List_WorkType")
End Sub

%>