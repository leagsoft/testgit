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
'//保存修改的资源记录
Sub SaveMdy()

    Dim Id,ClassId,Title,Url,Author,From,KeyWord
    Dim Editor,Count,Speciality
    Dim Content,ImgNews,SmallImg,BigImg,ShortContent,NowTime
    Dim CreateFile
    '取得参数
    Id=CLng(Request("Id"))
    ClassId=Request("radioBoxItem")
    If ClassId="" Then
        Response.Write("<script>alert(""请设置[资源类别]"");window.history.back();</script>")
        Response.End
    End If
    '**** Add By BennyLIu:20040625   '定义浏览者
    '**If Session("QXMC")="金融统计信息" then
	If Session("QXMC")="金融统计信息" or Session("QXMC")="分局动态" then		'Modify By BennyLiu:20040712
		Browser=Request("Browser")					'取得浏览者				
		DocumentType=Trim(Request("DocumentType")) '取得文件类型
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
    '//入库
    Dim Sql
    Sql="Select Top 1 * From News Where Id=" & ID & " Order By ID DESC"
    Dim Rs
    Set Rs=Server.CreateObject("ADODB.RecordSet")    
    Rs.Open Sql,Conn,1,3
    If Rs.Eof And Rs.Bof Then
        Rs.Close
        Set Rs=Nothing
        Response.Write("<script>alert(""记录不存在"");window.history.back();</script>")
        Response.End
    End If
    Rs("Class")=ClassId
    Rs("Title")=Title
    Rs("Url")=Url
    'Modify By BennyLIU:20040625
    
    if Session("QXMC")="局长专题" then  
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

    '完成
    Response.Redirect("News_List.asp?Work="&Request.Cookies("ZGW_NewsSys")("News_List_WorkType"))
End Sub


'///////////////////////////
'//添加资源记录
Sub AddReco()
    Dim ClassId,Title,Url,Author,From,KeyWord
    Dim Editor,Count
    Dim Content,ImgNews,SmallImg,BigImg,ShortContent,NowTime,UserBM
    '取得参数
    ClassId=Request("radioBoxItem")
    If ClassId="" Then
        Response.Write("<script>alert(""请设置[资源类别]"");window.history.back();</script>")
        Response.End
    End If
    'Add By BennyLiu:20040628，为了定义浏览者权限而增加
	'**If session("QXMC")="金融统计信息" then
	If session("QXMC")="金融统计信息" or Session("QXMC")="分局动态" then		'Modify By BennyLiu:20040712
		Browser=Request("Browser")		'取得可浏览者
		DocumentType=Trim(Request("DocumentType"))	'取得文件类型
	End If
	'End Add
    Title=Request("Title")
    Url=Request("Url")
    Author=Request("Author")		'资源作者，保存局长的序号
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
    '入库
    Rs("Class")=ClassId
    Rs("Title")=Title
    Rs("Url")=Url
    'Modify By BennyLiu:20040625
    If Session("QXMC")="局长专题" then
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
	If Browser<>"" then		'浏览者进库
		Rs("Browser")=Browser
	End If
	If DocumentType<>"" then
		Rs("IsDocument")=DocumentType		'文件类型进库
	End If
    Rs.Update
    Rs.Close
    Set Rs=Nothing
    '定义一个Session，方便下次录入信息时不用再输入相同的浏览者
    If Browser<>"" then
		Session("Browser")=Browser
		Session.Timeout =30
    End If
    Response.Write("<script>if(confirm(""<添加成功>\n是否想继续添加资源？"")){window.location='News_Add.asp?Work=AddReco'}else{window.location='"&"News_List.asp?Work="&Request.Cookies("ZGW_NewsSys")("News_List_WorkType")&"'}</script>")
	If CBool(Request("CurrentClassIdUsed")) Then
		Response.Cookies("ZGW_NewsSys")("CurrentClassIdUsed")=ClassId
	Else
		Response.Cookies("ZGW_NewsSys")("CurrentClassIdUsed")=""
	End If
End Sub

'//////////////////
'//真实,虚拟删除
Sub DelReco()
    '当前管理员分类权限等级值
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
        '是否具备足够的权限
        'If ClassPopedomType=SysAdmin.defClassPopedomType_Hig Then
            Sql="Select * From News Where Id=" & CLng(IdList(I))
            Rs.Open Sql,Conn,1,3
            If Not(Rs.Eof And Rs.Bof) Then
                '删除物理文件
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
    LogClass.AddLog(SysAdmin.AdminTitle & "删除"&UBound(IdList)+1&"条资源,实际成功"&FinishedNum&"条")
    Set LogClass=Nothing
    Response.Write("<script>alert(""<系统提示>\n此次操作预计操作（"&UBound(IdList)+1&"）条资料,实际成功操作（"&FinishedNum&"）条，未完成（"&(UBound(IdList)+1-FinishedNum)&"）条"");window.location="""&"News_List.asp?Work="&Request.Cookies("ZGW_NewsSys")("News_List_WorkType")&"&CurrentPage="&Request.Cookies("ZGW_NewsSys")("News_List_CurrentPage")&""";</script>")
    Response.End
End Sub


'/////////////////
'//清空回收站
Sub ClearDustbin()
    
    Dim Sql
        Sql="DELETE FROM News WHERE DEL=1"
    Conn.ExeCute(Sql)

    Dim LogClass
    Set LogClass=New Tkl_LogClass
    LogClass.AddLog(SysAdmin.AdminTitle & "清空回收站")
    Set LogClass=Nothing

    Response.Redirect "News_List.asp?Work="&Request.Cookies("ZGW_NewsSys")("News_List_WorkType")
End Sub

%>