<%
'--------------------------------------------------------------------
'函数名称：DateToString(时间)
'功能：日期转换
Function DateToString(dDate)
    DateToString = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)
End Function

'--------------------------------------------------------------------
'函数名称：DateToString(时间)
'功能：时间转换
Function TimeToString(tTime)
    TimeToString = RIGHT("00" + Trim(Hour(tTime)),2) + ":" + RIGHT("00"+Trim(Minute(tTime)),2) + ":" + RIGHT("00"+Trim(Second(tTime)),2)
End Function

'----------------------------------------------------------
'函数名称：GetDic(下拉菜单名称,需选择的值,语言,字典类型)
'功能：根据表名及字段名从字典表中查出相应的值，并生成一个下拉菜单
Function GetDic(cSelectName,cSelected,cLang,cType)
     If IsNull(cLang) or IsEmpty(cLang) Then cLang=1
     theDic = vbCrlf&"<select name="&cSelectName&">"&vbCrLf&_
              "<option value=""""> 请 选 择 </option>"&vbCrLf
     '查询数据库
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
'函数名称：GetCheckBox(数组,CheckBox名称,选择的值)
'功能：根据值的数组、CheckBox名称及需选择的值生成一组CheckBox
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
'函数名称：PutEvent(事件类型,用户帐号,事件内容,是否调用系统信息页,返回的URL，继续的URL)
'功能：将系统事件写进Event数据库
Function PutEvent(dType,dUserId,dMsg,dRedirect,dBack,dContinue)
     '接收用户IP及浏览器信息
     dIp      = Trim(Request.ServerVariables("REMOTE_ADDR"))
     dBrowser = Trim(Request.ServerVariables("HTTP_USER_AGENT"))
     If dRedirect = "" Then dRedirect = "N"
     
     '设定日志文件路径
     vPhysicalPath = Server.MapPath("/LogFiles/Event"+DateToString(Now())+".log")
     
     '创建文件系统对象
     Set fs = CreateObject("Scripting.FileSystemObject")
        
     '追加事件到日志文件中
     Set fo = fs.OpenTextFile(vPhysicalPath,8,true)
     fo.WriteLine(Now()&"  "&dType&"  用户"&dUserId&"(来自"&dIp&"):"&dMsg)
     fo.Close     
     Set fs = Nothing

     '调用系统信息页
     If dRedirect = "Y" Then
        Response.Redirect "/Event/Index.asp?cMsg="&Server.UrlEncode(dMsg)&"&cBack="&dBack&"&cContinue="&dContinue
     End If
End Function  

'----------------------------------------------------------
'函数名称：PutEvent(事件类型,用户帐号,事件内容,是否调用系统信息页,返回的URL，继续的URL)
'功能：将系统事件写进Event数据库
Function GetEvent(dType,dUserId,dMsg,dRedirect,dBack,dContinue)
     '接收用户IP及浏览器信息
     dIp      = Trim(Request.ServerVariables("REMOTE_ADDR"))
     dBrowser = Trim(Request.ServerVariables("HTTP_USER_AGENT"))
     If dRedirect = "" Then dRedirect = "N"
     
     '设定日志文件路径
     vPhysicalPath = Server.MapPath("/LogFiles/Event"+DateToString(Now())+".log")
     
     '创建文件系统对象
     Set fs = CreateObject("Scripting.FileSystemObject")
        
     '追加事件到日志文件中
     Set fo = fs.OpenTextFile(vPhysicalPath,8,true)
     fo.WriteLine(Now()&"  "&dType&"  用户"&dUserId&"(来自"&dIp&"):"&dMsg)
     fo.Close     
     Set fs = Nothing

     '调用系统信息页
     If dRedirect = "Y" Then
        Response.Redirect "/Event/Index.asp?cMsg="&Server.UrlEncode(dMsg)&"&cBack="&dBack&"&cContinue="&dContinue
     End If
End Function  

'--------------------------------------------------------------------
'函数名称：GetPruductEof(条件语句)
'功能：查询某一产品分类下是否有产品，并返回整型(0为没有，1为有)
Function GetProductEof(dSql)
     '根据用户权限生成不同的查询语句
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
'函数名称：ShowBody(文本)
'功能：格式化文本
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
'需询盘的值
dim arrFancy(22)
arrFancy(0) = "逛商城"
arrFancy(1) = "艺术"
arrFancy(2) = "计算机/网"
arrFancy(3) = "电脑游戏"
arrFancy(4) = "旅行"
arrFancy(5) = "储蓄/投资 "
arrFancy(6) = "烹调"
arrFancy(7) = "饮酒/品茶 "
arrFancy(8) = "社交"
arrFancy(9) = "进修"
arrFancy(10) = "电影/电视"
arrFancy(11) = "音乐"
arrFancy(12) = "阅读"
arrFancy(13) = "购物"
arrFancy(14) = "法律"
arrFancy(15) = "高尔夫球"
arrFancy(16) = "健身/跑步"
arrFancy(17) = "钓鱼"
arrFancy(18) = "游泳"
arrFancy(19) = "滑冰/滑雪"
arrFancy(20) = "网球"
arrFancy(21) = "乒乓球"
arrFancy(22) = "骑单车/摩托车"   
%>