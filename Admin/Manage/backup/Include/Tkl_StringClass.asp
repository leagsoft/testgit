<%
'////////////////////////////////////////////////////////
'//                    �ַ����������
'///////////////////////////////////////////////////////

Class Tkl_StringClass
    '//��Html��ǩ��ȡ���ı�����
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

    '//���Email
    '//����:True/False
    Public Function CheckEmail(strng)
        CheckEmail = false
        Dim regEx, Match
        Set regEx = New RegExp
        regEx.Pattern = "^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$"
        regEx.IgnoreCase = True
        Set Match = regEx.Execute(strng)
        if match.count then CheckEmail= true
    End Function

    '//�ַ����Ƿ���[0-9]&[a-z]���»����У������ִ�Сд��
    '//����:True/False
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

    '//�ַ����Ƿ���[a-z]�У������ִ�Сд��
    '//����:True/False
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

    '//�ַ����Ƿ���[0-9]�У������ִ�Сд��
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

    '//Html�ַ���תJs�ַ���
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

    '//ɨ��Ԫ��mItem�Ƿ���Ԫ���б�strItemList��
    '//����:stritemList(��ɨ��Ԫ���б���Ԫ���Զ��Ÿ���),mItem(��ƥ��Ԫ��)
    '//����:True/False
    '//����myCharClass.ItemInList("1,2,3,34,23","2",",") ���:True
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

    '//ת��Html�ؼ���ǩΪHtml�����ַ���
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

    '//ת��Html�ؼ���ǩΪHtml�����ַ���(��ת��Ӳ�س�����س���)
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

    '//�������ַ����滻
    '//������������ʽ,���滻�ַ���,�滻�ַ���
    Public Function ReplaceTest(patrn,mStr,replStr)
        Dim regEx
        Set regEx = New RegExp
        regEx.Pattern = patrn
        regEx.IgnoreCase = True
        regEx.Global = True
        ReplaceTest = regEx.Replace(mStr,replStr)
    End Function

    '//�������ַ�������
    '//������������ʽ,���滻�ַ���,�滻�ַ���
    '//���أ�Bool(True:�ҵ�)
    Public Function FindText(patrn,mStr)
        Dim regEx
        Set regEx = New RegExp
        regEx.Pattern = patrn
        regEx.IgnoreCase = True
        regEx.Global = True
        FindText = regEx.Test(mStr)
    End Function

    '//����Ƿ��н�ֹ�ַ���
    '//����:������ַ���,��ֹ�ַ��б�(��,�Ÿ���)
    '//����:True(����Υ���ַ�)/False
    '//����myCharClass.BadWord("����������˵���Fuck You","fuck you,���˵�,you are pig")
    Public Function BadWord(str,BadWordList)
        BadWord=False
        Dim arrBadWord
            arrBadWord=Split(BadWordList,",",-1,1)
        Dim regEx
        Set regEx=New RegExp
        regEx.IgnoreCase = True            '�����ִ�Сд
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
    
    '//��ȡָ�������ַ���
    '//�������ͣ��ַ���
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
    
    '//ʱ���ʽ��
    '//������ʱ�䣬��ʽģ��
    '//���أ���ʽ������ַ���
    '//��ע����ʽ���ؼ�����⣺
    '       "{Y}" : 4λ��
    '       "{y}" : 2λ��
    '       "{M}" : ����λ����
    '       "{m}" : ��λ����,��03,01
    '       "{D}" : ����λ����
    '       "{d}" : ��λ����
    '       "{H}" : ����λ��Сʱ
    '       "{h}" : ��λ��Сʱ
    '       "{MI}": ����λ�ķ���
    '       "{mi}": ��λ�ķ���
    '       "{S}" : ����λ����
    '       "{s}" : ��λ����
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