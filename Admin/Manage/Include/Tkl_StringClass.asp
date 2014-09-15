<%
'////////////////////////////////////////////////////////
'//                    字符串函数类库
'///////////////////////////////////////////////////////

Class Tkl_StringClass
    '//从Html标签中取出文本内容
    Public Function GetTextFromHtml(strHtml)
        Dim strPatrn
            strpatrn="<.*?>"
        Dim regEx
        Set regEx = New RegExp
        regEx.Pattern = strPatrn
        regEx.IgnoreCase = True
        regEx.Global = True
        GetTextFromHtml = regEx.Replace(strHtml,"")
    End Function

    '//检测Email
    '//返回:True/False
    Public Function CheckEmail(strng)
        CheckEmail = false
        Dim regEx, Match
        Set regEx = New RegExp
        regEx.Pattern = "^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$"
        regEx.IgnoreCase = True
        Set Match = regEx.Execute(strng)
        if match.count then CheckEmail= true
    End Function

    '//字符串是否在[0-9]&[a-z]及下划线中（不区分大小写）
    '//返回:True/False
    Public Function IsChar26AndInt(str)
        IsChar26AndInt=True
        Dim regEx,Match
        Set regEx=New RegExp
            regEx.Pattern="[\W]{1,}?"
            regEx.IgnoreCase=True
        Set Match=regEx.Execute(str)
        If Match.Count>=1 Then
            IsChar26AndInt=False
        End If
    End Function

    '//字符串是否在[a-z]中（不区分大小写）
    '//返回:True/False
    Public Function IsChar26(str)
        IsChar26=True
        Dim regEx,Match
        Set regEx=New RegExp
            regEx.Pattern="[^a-zA-Z]{1,}?"
            regEx.IgnoreCase=True
        Set Match=regEx.Execute(str)
        If Match.Count>=1 Then
            IsChar26=False
        End If
    End Function

    '//字符串是否在[0-9]中（不区分大小写）
    Public Function IsIntChar(str)
        IsIntChar=True
        Dim regEx,Match
        Set regEx=New RegExp
            regEx.Pattern="\D{1,}?"
            regEx.IgnoreCase=True
        Set Match=regEx.Execute(str)
        If Match.Count>=1 Then
            IsIntChar=False
        End If
    End Function

    '//Html字符串转Js字符串
    Public Function HTMLToJS(strHtml)
        If Trim(strHtml)="" Then
            HTMLToJS=""
            Exit Function
        End If
        strHtml=Replace(strHtml,"\","\\")
        strHtml=Replace(strHtml,"""","\""")
        strHtml=Replace(strHtml,vbCrLf,"")
        HTMLToJS=strHtml
    End Function

    '//扫描元素mItem是否在元素列表strItemList中
    '//参数:stritemList(被扫描元素列表，各元素以逗号隔开),mItem(欲匹配元素)
    '//返回:True/False
    '//例：myCharClass.ItemInList("1,2,3,34,23","2",",") 结果:True
    Public Function ItemInList(strItemList,mItem)
        ItemInList=False
        If IsNull(strItemList) Or IsNull(mItem="") Then
            Exit Function
        End If
        strItemList=Replace(strItemList," ","")
        If Instr(","&strItemList&",",","&mItem&",")>=1 Then
            ItemInList=True
        End If
    End Function

    '//转换Html关键标签为Html特殊字符串
    Public Function HTMLEncode(str)
        If Not Isnull(str) Then
            str = Replace(str, CHR(13), "")
            str = Replace(str, CHR(10) & CHR(10), "<P></P>")
            str = Replace(str, CHR(10), "<BR>")
            str = replace(str, ">", "&gt;")
            str = replace(str, "<", "&lt;")
            str = replace(str, "&",    "&amp;")
            str = replace(str, " ",    "&nbsp;")
            str = replace(str, """", "&quot;")
            HTMLEncode = str
        End If
    End Function

    '//转换Html关键标签为Html特殊字符串(不转换硬回车及软回车符)
    Public Function HTMLEncode2(str)
        If Not Isnull(str) Then
            str = replace(str, ">", "&gt;")
            str = replace(str, "<", "&lt;")
'            str = replace(str, "&",    "&amp;")
'            str = replace(str, " ",    "&nbsp;")
'            str = replace(str, """", "&quot;")
            HTMLEncode2 = str
        End If
    End Function

    '//函数：字符串替换
    '//参数：正则表达式,被替换字符串,替换字符串
    Public Function ReplaceTest(patrn,mStr,replStr)
        Dim regEx
        Set regEx = New RegExp
        regEx.Pattern = patrn
        regEx.IgnoreCase = True
        regEx.Global = True
        ReplaceTest = regEx.Replace(mStr,replStr)
    End Function

    '//函数：字符串查找
    '//参数：正则表达式,被替换字符串,替换字符串
    '//返回：Bool(True:找到)
    Public Function FindText(patrn,mStr)
        Dim regEx
        Set regEx = New RegExp
        regEx.Pattern = patrn
        regEx.IgnoreCase = True
        regEx.Global = True
        FindText = regEx.Test(mStr)
    End Function

    '//检测是否含有禁止字符串
    '//参数:被检测字符串,禁止字符列表(以,号隔开)
    '//返回:True(含有违禁字符)/False
    '//例：myCharClass.BadWord("你他妈的王八蛋，Fuck You","fuck you,王八蛋,you are pig")
    Public Function BadWord(str,BadWordList)
        BadWord=False
        Dim arrBadWord
            arrBadWord=Split(BadWordList,",",-1,1)
        Dim regEx
        Set regEx=New RegExp
        regEx.IgnoreCase = True            '不区分大小写
        regEx.Global = True
        Dim Match
        Dim I
        For I=0 To UBound(arrBadWord)
            response.write arrBadWord(I)&"<br>"
            If arrBadWord(I)<>"" Then
                regEx.Pattern=arrBadWord(I)
                Set Match=regEx.Execute(str)
                If Match.Count Then
                    BadWord=True
                    Exit For
                End If
            End If
        Next
    End Function
    
    '//截取指定长度字符串
    '//返回类型：字符串
    Public Function CutStr(str,strlen)
        dim l,t,c,m_i
        l=len(str)
        t=0
        for m_i=1 to l
            c=Abs(Asc(Mid(str,m_i,1)))    
            if c>255 then
                t=t+2
            else
                t=t+1
            end if
    
            if t>=strlen then
                CutStr=left(str,m_i)&"..."
                exit for
            else
                CutStr=str
            end if
        next
    End Function
    
    '//时间格式化
    '//参数：时间，格式模板
    '//返回：格式化后的字符串
    '//备注：格式化关键词详解：
    '       "{Y}" : 4位年
    '       "{y}" : 2位年
    '       "{M}" : 不补位的月
    '       "{m}" : 补位的月,如03,01
    '       "{D}" : 不补位的日
    '       "{d}" : 补位的日
    '       "{H}" : 不补位的小时
    '       "{h}" : 补位的小时
    '       "{MI}": 不补位的分钟
    '       "{mi}": 补位的分钟
    '       "{S}" : 不补位的秒
    '       "{s}" : 补位的秒
    Public Function FormatMyDate(myDate,Template)
        If Not IsDate(myDate) Or Template = "" Then
            FormatMyDate = Template
            Exit Function
        End If

        Dim mYear,mMonth,mDay,mHour,mMin,mSec
            mYear = Year(myDate)
            mMonth = Month(myDate)
            mDay = Day(myDate)
            mHour = Hour(myDate)
            mMin = Minute(myDate)
            mSec = Second(myDate)

        Template = Replace(Template,"{Y}",Year(myDate))
        Template = Replace(Template,"{y}",Right(Year(myDate),2))
        Template = Replace(Template,"{M}",Month(myDate))
        Template = Replace(Template,"{m}",Right("00" & Month(myDate),2))
        Template = Replace(Template,"{D}",Day(myDate))
        Template = Replace(Template,"{d}",Right("00" & Day(myDate),2))
        Template = Replace(Template,"{H}",Hour(myDate))
        Template = Replace(Template,"{h}",Right("00" & Day(myDate),2))
        Template = Replace(Template,"{MI}",Minute(myDate))
        Template = Replace(Template,"{mi}",Right("00" & Minute(myDate),2))
        Template = Replace(Template,"{S}",Second(myDate))
        Template = Replace(Template,"{s}",Right("00" & Second(myDate),2))

        FormatMyDate = Template
    End Function
End Class
%>