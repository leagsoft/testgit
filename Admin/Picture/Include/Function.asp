<%
'--------------------------------------------------------------------
'�������ƣ�DateToString(ʱ��)
'���ܣ�����ת��
Function DateToString(dDate)
    DateToString = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)
End Function

'--------------------------------------------------------------------
'�������ƣ�DateToString(ʱ��)
'���ܣ�ʱ��ת��
Function TimeToString(tTime)
    TimeToString = RIGHT("00" + Trim(Hour(tTime)),2) + ":" + RIGHT("00"+Trim(Minute(tTime)),2) + ":" + RIGHT("00"+Trim(Second(tTime)),2)
End Function

'----------------------------------------------------------
'�������ƣ�GetDic(�����˵�����,��ѡ���ֵ,����,�ֵ�����)
'���ܣ����ݱ������ֶ������ֵ���в����Ӧ��ֵ��������һ�������˵�
Function GetDic(cSelectName,cSelected,cLang,cType)
     If IsNull(cLang) or IsEmpty(cLang) Then cLang=1
     theDic = vbCrlf&"<select name="&cSelectName&">"&vbCrLf&_
              "<option value=""""> �� ѡ �� </option>"&vbCrLf
     '��ѯ���ݿ�
     cSql = "select value from sysdic where lang="&cLang&" and type='"&cType&"' order by dicid asc"
     Set cRs = mydb.Execute(cSql)
     Do
       If cRs.Eof Then Exit Do
          theDic = theDic&"<option value='"&Trim(cRs("value"))&"' "
          If cSelected=Trim(cRs("value")) Then theDic=theDic&"selected"
          theDic = theDic&">"&Trim(cRs("value"))&"</option>"&vbCrLf
       cRs.MoveNext
     Loop
     cRs.Close
     Set cRs = Nothing
     theDic = theDic&"</select>"
     GetDic = theDic 
End Function

'----------------------------------------------------------
'�������ƣ�GetCheckBox(����,CheckBox����,ѡ���ֵ)
'���ܣ�����ֵ�����顢CheckBox���Ƽ���ѡ���ֵ����һ��CheckBox
Function GetCheckBox(arr,name,value)
	dim str
	str = "<table width='96%' border='0' cellspacing='0' cellpadding='0' align='center'><tr>"
	for i=0 to UBound(arr)
		if InStr(value,arr(i))>0 then
		   str = str & "<td><input type='checkbox' checked name='" _
		       & name & "' value='" & arr(i) & "'>&nbsp;" & arr(i) & "</td>"
		else
		   str = str & "<td><input type='checkbox' name='" _
		       & name & "' value='" & arr(i) & "'>&nbsp;" & arr(i) & "</td>"
		end if
		if ((i+1) Mod 3)=0 then
		   str = str & "</tr><tr>"
		end if
	next
	str = str & "</tr></table>"
	GetCheckBox = str
End Function

'----------------------------------------------------------
'�������ƣ�PutEvent(�¼�����,�û��ʺ�,�¼�����,�Ƿ����ϵͳ��Ϣҳ,���ص�URL��������URL)
'���ܣ���ϵͳ�¼�д��Event���ݿ�
Function PutEvent(dType,dUserId,dMsg,dRedirect,dBack,dContinue)
     '�����û�IP���������Ϣ
     dIp      = Trim(Request.ServerVariables("REMOTE_ADDR"))
     dBrowser = Trim(Request.ServerVariables("HTTP_USER_AGENT"))
     If dRedirect = "" Then dRedirect = "N"
     
     '�趨��־�ļ�·��
     vPhysicalPath = Server.MapPath("/LogFiles/Event"+DateToString(Now())+".log")
     
     '�����ļ�ϵͳ����
     Set fs = CreateObject("Scripting.FileSystemObject")
        
     '׷���¼�����־�ļ���
     Set fo = fs.OpenTextFile(vPhysicalPath,8,true)
     fo.WriteLine(Now()&"  "&dType&"  �û�"&dUserId&"(����"&dIp&"):"&dMsg)
     fo.Close     
     Set fs = Nothing

     '����ϵͳ��Ϣҳ
     If dRedirect = "Y" Then
        Response.Redirect "/Event/Index.asp?cMsg="&Server.UrlEncode(dMsg)&"&cBack="&dBack&"&cContinue="&dContinue
     End If
End Function  

'----------------------------------------------------------
'�������ƣ�PutEvent(�¼�����,�û��ʺ�,�¼�����,�Ƿ����ϵͳ��Ϣҳ,���ص�URL��������URL)
'���ܣ���ϵͳ�¼�д��Event���ݿ�
Function GetEvent(dType,dUserId,dMsg,dRedirect,dBack,dContinue)
     '�����û�IP���������Ϣ
     dIp      = Trim(Request.ServerVariables("REMOTE_ADDR"))
     dBrowser = Trim(Request.ServerVariables("HTTP_USER_AGENT"))
     If dRedirect = "" Then dRedirect = "N"
     
     '�趨��־�ļ�·��
     vPhysicalPath = Server.MapPath("/LogFiles/Event"+DateToString(Now())+".log")
     
     '�����ļ�ϵͳ����
     Set fs = CreateObject("Scripting.FileSystemObject")
        
     '׷���¼�����־�ļ���
     Set fo = fs.OpenTextFile(vPhysicalPath,8,true)
     fo.WriteLine(Now()&"  "&dType&"  �û�"&dUserId&"(����"&dIp&"):"&dMsg)
     fo.Close     
     Set fs = Nothing

     '����ϵͳ��Ϣҳ
     If dRedirect = "Y" Then
        Response.Redirect "/Event/Index.asp?cMsg="&Server.UrlEncode(dMsg)&"&cBack="&dBack&"&cContinue="&dContinue
     End If
End Function  

'--------------------------------------------------------------------
'�������ƣ�GetPruductEof(�������)
'���ܣ���ѯĳһ��Ʒ�������Ƿ��в�Ʒ������������(0Ϊû�У�1Ϊ��)
Function GetProductEof(dSql)
     '�����û�Ȩ�����ɲ�ͬ�Ĳ�ѯ���
     Select Case Session("cAllowSys")
            Case 0
                 cSql = "select PROID from PRODUCTS where DELETED=0"&dSql&" and ALLOWSYS=0"
            Case Else
                 cSql = "select PROID from PRODUCTS where DELETED=0"&dSql
     End Select 
     Set dRs = mydb.Execute(cSql)
     If dRs.Eof Then
        GetProductEof = 0
     Else
        GetProductEof = 1
     End If
     dRs.Close
     Set dRs = Nothing
End Function

'--------------------------------------------------------------------
'�������ƣ�ShowBody(�ı�)
'���ܣ���ʽ���ı�
Function ShowBody(Str)
     dim dist
     dim i
     If Not IsNull(Str) or IsEmpty(Str) or Str="" Then 
        For i = 1 to Len(Str)
            If mid(Str,i,1)<>"%" and ucase(mid(Str,i,6))<>"SCRIPT" then
               If mid(str,i,1)<>chr(13) then
                  dist=dist+mid(Str,i,1)
               Else
	          response.write dist
                  response.write "<BR>"+chr(13)+chr(10)
	          dist=""
               End If
            End If
        Next
        ShowBody=dist
     End If 
End Function

'--------------------------------------------------------------------
'��ѯ�̵�ֵ
dim arrFancy(22)
arrFancy(0) = "���̳�"
arrFancy(1) = "����"
arrFancy(2) = "�����/��"
arrFancy(3) = "������Ϸ"
arrFancy(4) = "����"
arrFancy(5) = "����/Ͷ�� "
arrFancy(6) = "���"
arrFancy(7) = "����/Ʒ�� "
arrFancy(8) = "�罻"
arrFancy(9) = "����"
arrFancy(10) = "��Ӱ/����"
arrFancy(11) = "����"
arrFancy(12) = "�Ķ�"
arrFancy(13) = "����"
arrFancy(14) = "����"
arrFancy(15) = "�߶�����"
arrFancy(16) = "����/�ܲ�"
arrFancy(17) = "����"
arrFancy(18) = "��Ӿ"
arrFancy(19) = "����/��ѩ"
arrFancy(20) = "����"
arrFancy(21) = "ƹ����"
arrFancy(22) = "�ﵥ��/Ħ�г�"   
%>