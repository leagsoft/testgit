<%
'---------------------------------------------------------------
  page=Request("page")  
'Connect2="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=NEWSCBRCGD;Data Source=oa-server2;Pwd=weboaadmin2004"
Connect2="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=PBC;Data Source=192.168.0.201;Pwd="
  Set CNMain2=Server.CreateObject("ADODB.Connection")
  CNMain2.open Connect2,"",""
' ---------------------------------------------------------------

' ��Domino����������֤�û��Ƿ�Ϸ�
function DominoLogin(Username, Password)
    
    'const DOMINO_LOGIN_URL = "http://oademo/ahz.nsf/LOGIN?OpenAgent"
    const DOMINO_LOGIN_URL = "http://rh28xx07/ahz.nsf/LOGIN?OpenAgent"
    const DOMINO_LOGIN_OK  = "DOMINO_LOGIN_OK"
    
    dim obj, Text
    Set obj = CreateObject("UrlCall.Api")
    Text = obj.DownloadUrl(DOMINO_LOGIN_URL, Username, Password)
    obj = null
    DominoLogin = InStr(Text, DOMINO_LOGIN_OK)>0

    if Text="" then
        Response.Write "�޷���֤���˵�ַ���ɷ��ʣ�" + DOMINO_LOGIN_URL + "<br><br>"
    end if
    'Response.Write Text
end function

function FieldGet(byval mystr,byval Pos,byval delimiter)
    '---------------------------------------------------------------
    '// ����ֲ�����
    Dim count,leng,index,pindex
    '---------------------------------------------------------------
    '// �������ַ���Ϊ��,��������С��1,�򷵻ؿ��ַ���
    If (mystr = "") Or (delimiter = "") Or (Pos < 1) Then
       FieldGet = ""
       Exit Function
       End If
    '---------------------------------------------------------------
    '// ѭ������,��ָ��ָ��Ҫ��ȡ�����ʼλ��
    count = 1
    pindex=1
    leng=len(delimiter)
    index=InStr(1,myStr, delimiter)
    While (index>0)and(count<pos)
       count = count + 1
       pindex = index + leng
       index = InStr(index + leng, mystr, delimiter)
       Wend
    '---------------------------------------------------------------
    If (count < pos) Then
       FieldGet = ""
       Else
       If (index > 0) Then FieldGet = Mid(mystr, pindex, index - pindex)
       If (index <= 0) Then FieldGet = Mid(mystr, pindex, Len(mystr) - pindex + 1)
       End If
    '---------------------------------------------------------------
end function
function FieldCount(byval mystr,byval delimiter)
    '---------------------------------------------------------------
    '//����ֲ�����
    Dim count,leng,index
    '---------------------------------------------------------------
    '// �������ַ���Ϊ��,�򷵻�����ĿΪ1
    If (mystr= "") Or (delimiter ="") Then
       FieldCount = 1
       Exit Function
       End If
    '---------------------------------------------------------------
    '// ѭ������,��ָ��ָ��Ҫ��ȡ�����ʼλ��
    count = 1
    leng=len(delimiter)
    index=InStr(1,myStr, delimiter)
    While index > 0
       count = count + 1
       index=InStr(index+leng,myStr, delimiter)
       Wend
    FieldCount = count
    '---------------------------------------------------------------
end function
function FieldGet_(byval str,byval idx,byval ch)
    '---------------------------------------------------------------
    If(str="")or(ch="")or(idx<1)Then
       FieldGet = ""
       else
       temp=split(str,ch)
       if(idx>ubound(temp)+1)then
          FieldGet = ""
          else
          FieldGet=temp(idx-1)
          end if
       end If
    '---------------------------------------------------------------
end function
function selectdate(byval date1,byval date2,byval date3,byval date4)
    '---------------------------------------------------------------
    '//��date1->date2->date3->date4˳��ѡ����ֵ����
    If not isnull(date1) or date1<>"" Then selectdate=date1
    If (isnull(date1) or date1="") and (not isnull(date2) or date2<>"") Then selectdate=date2
    If (isnull(date1) or date1="") and (isnull(date2) or date2="") and (not isnull(date3) or date3<>"") Then selectdate=date3
    If (isnull(date1) or date1="") and (isnull(date2) or date2="") and (isnull(date3) or date3="") and (not isnull(date4) or date4<>"") Then selectdate=date4      
    '---------------------------------------------------------------
end function
function FieldCount_(byval str,byval ch)
    '---------------------------------------------------------------
    If(str="")or(ch="")Then
       Fieldcount=1
       else
       temp=split(str,ch)
       Fieldcount=ubound(temp)+1
       end If
    '---------------------------------------------------------------
end function
function MyTime3()
  '// ��ָ���ĸ�ʽ��ʾ����
  '---------------------------------------------------------------
  dim yy,mm,dd
  '---------------------------------------------------------------
  yy = right("20" & Year(Date()),4)
  mm = right("00" & Month(Date()),2)
  dd = right("00" & Day(Date()),2)
   MyTime = yy+"-"+mm+"-"+dd
  '---------------------------------------------------------------
end function
function MyTime()
  '// ��ָ���ĸ�ʽ��ʾ����/ʱ��
  '---------------------------------------------------------------
  dim yy,mm,dd,hh,nn,ss
  '---------------------------------------------------------------
  yy = right("20" & Year(Date()),4)
  mm = right("00" & Month(Date()),2)
  dd = right("00" & Day(Date()),2)
  hh = right("00" & hour(time()),2)
  nn = right("00" & minute(time()),2)
  ss = right("00" & second(time()),2)
  MyTime = yy+"-"+mm+"-"+dd+" "+hh+":"+nn+":"+ss
  '---------------------------------------------------------------
end function
function MyTime2(byval thetime)
  '// ��ָ���ĸ�ʽ��ʾ����/ʱ��
  '---------------------------------------------------------------
  dim yy,mm,dd,hh,nn,ss
  '---------------------------------------------------------------
  yy = right("20" & Year(thetime),4)
  mm = right("00" & Month(thetime),2)
  dd = right("00" & Day(thetime),2)
  hh = right("00" & hour(thetime),2)
  nn = right("00" & minute(thetime),2)
  ss = right("00" & second(thetime),2)
  MyTime2 = yy+"-"+mm+"-"+dd+" "+hh+":"+nn+":"+ss
  '---------------------------------------------------------------
end function

function format8date(byval pddate)
  '// ��ʾyyyymmdd���ڸ�ʽ
		dim strYear,strMonth,strDay
		Dim strOutPutDate
		IF isdate(pddate)  then
			strYear=year(pddate)
			strMonth=month(pddate)
			strDay=day(pddate)
			
			strOutPutDay=strYear & Right("0" & strMonth,2) & right("0" & strDay,2)
		else
			strOutPutDay=""
		end if
		
		format8date=strOutPutDay
	

end function
function GetCommand(byval table,byval key,byval flist,byval cond2,byval order)
  '// ���ݲ������ָ����SQL��ѯ���
  '// ����û��������Ҫ��ѯ��[�ؼ���]����õ���Ӧ�Ĳ�ѯ���
  '---------------------------------------------------------------
  '// table:����
  '// key  :��Ҫ�����Ĺؼ���
  '// flist:��Ҫ�����ؼ��ֵ��ֶ�
  '// cond2:���Ӳ�ѯ����
  '---------------------------------------------------------------
  dim ii,cond,temp,temp2
  '---------------------------------------------------------------
  '// ����ؼ���Ϊ��,����ݷ��ز�ѯ���
  if(key="")then
     cond="select * from "+table+" "+order
     if(cond2<>"")then cond="select * from "+table+" where ("+cond2+") "+order
     GetCommand=cond
     exit function
     end if
  '---------------------------------------------------------------
  '// ����û��������[��ѯ����]����õ���Ӧ�Ĳ�ѯ���
  temp2=left(lcase(trim(key)),6)
  '// ����where�����õ���Ӧ�Ĳ�ѯ���
  if(temp2="where ")then
     temp2=replace(trim(key),"""","'")
     cond="select * from "+table+" where ("+right(temp2,len(temp2)-6)+")"
     if(cond2<>"")then cond=cond+"and("+cond2+")"
     GetCommand=cond+" "+order
     exit function
     end if
  '// ����order�����õ���Ӧ�Ĳ�ѯ���
  if(temp2="order ")then
     temp2=trim(key)
     cond="select * from "+table+" order by "+right(temp2,len(temp2)-9)
     if(cond2<>"")then cond="select * from "+table+" where ("+cond2+") "+key
     GetCommand=cond
     exit function
     end if
  '// ����top�����õ���Ӧ�Ĳ�ѯ���
  if(left(temp2,4)="top ")then
     cond="select "+key+" * from "+table
     if(cond2<>"")then cond="select "+key+" * from "+table+" where ("+cond2+")"
     GetCommand=cond+" "+order
     exit function
     end if
  '---------------------------------------------------------------
  '// ����ؼ��ֲ�Ϊ��,�������Ҫ�������ֶ�����
  flist=replace(flist,",","+'~'+")
  '---------------------------------------------------------------
  '// �õ����ؼ��ֵĲ�ѯ����
  cond=" where (charindex('"+key+"',"+flist+")>0)"
  '// �õ�[��]�����Ķ�ؼ��ּ������
  if(instr(key,"|")>0)then
     temp=split(key,"|")
     cond="(charindex('"+temp(0)+"',"+flist+")>0)"
     for ii=1 to ubound(temp)
         cond=cond+"or(charindex('"+temp(ii)+"',"+flist+")>0)"
         next
     cond=" where ("+cond+")"
     end if
  '// �õ�[��]�����Ķ�ؼ��ּ������
  if(instr(key,"/")>0)then
     temp=split(key,"/")
     cond="(charindex('"+temp(0)+"',"+flist+")>0)"
     for ii=1 to ubound(temp)
         cond=cond+"and(charindex('"+temp(ii)+"',"+flist+")>0)"
         next
     cond=" where ("+cond+")"
     end if
  '---------------------------------------------------------------
  '// ����cond�����ϳ������Ĳ�ѯ���
  GetCommand="select * from "+table+cond+" "+order
  if(cond2<>"")then GetCommand="select * from "+table+cond+" and ("+cond2+") "+order
  '---------------------------------------------------------------
end function
sub Setbar(link,beginstring,PageIndex,PageCount,PageSize,RecordCount,endstring)
  '// ��ʾ[��һҳ][��һҳ]�ĵ�����
  '---------------------------------------------------------------
  bgcolor1="bgcolor=#FFFFFF" : bgcolor2="bgcolor=#000000"
  if(UserPZ4<>"��׼")then
     bgcolor1=" bgcolor=#546E86"
     bgcolor2=" bgcolor=#546E86"
     end if
  Response.Write("<center><table width=100% cellspacing=0 cellpadding=1 "+bgcolor1+"><tr height=20 "+bgcolor2+"><td>")
  'Response.Write("<center><table border=1 width=100% cellspacing=0 cellpadding=1 "+bgcolor1+"><tr valign=middle height=22"+bgcolor2+"><td>")
  Response.Write(beginstring)
  Response.Write("<font color=#ffffff>[��"& PageIndex & "/" & PageCount & "ҳ]")
  'Response.Write("<font color=#ffffff>[ÿҳ" & PageSize & "��/��" & RecordCount & "����¼]")
  if(PageIndex<=1)then Response.Write("<font color=#ffffff>[��ҳ][��һҳ]")
  if(PageIndex>1)then Response.Write("<a href="""+link+"&n=1""><font color=#ffffff>[��ҳ]</a><a href="""+link+"&n=" & PageIndex-1 & """><font color=#ffffff>[��һҳ]</a>")
  if(PageIndex>=PageCount)then Response.Write("<font color=#ffffff>[��һҳ][βҳ]")
  if(PageIndex<PageCount)then Response.Write("<a href="""+link+"&n=" & (PageIndex+1) & """><font color=#ffffff>[��һҳ]</a><a href="""+link+"&n=" & PageCount & """><font color=#ffffff>[βҳ]</a>")
  Response.Write(endstring)
  '---------------------------------------------------------------
end sub
sub alarm(byval msg)
  '// ���ؾ���/��ʾ��Ϣ
  '// stringhead��Html�ļ�ͷ��ȫ�ֱ���
  '---------------------------------------------------------------
  bgcolor1=" bgcolor=#A0C0E8"
  if(UserPZ4<>"��׼")then bgcolor1=" bgcolor=#FFFFE8"
  Response.write(stringhead)
  Response.write("<body "+bgcolor1+" leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>")
  Response.write("<center><font class=black>"+msg+"</font></center>")
  if(fieldget(session("sys.UserPZ"),8,chr(5))="��")then
     Response.write("<SCRIPT language=JavaScript><!--"+chr(13)+chr(10)+"alert('"+msg+"');"+chr(13)+chr(10)+"//--></SCRIPT>"+chr(13)+chr(10))
  end if
  Response.end
  '---------------------------------------------------------------
end sub

sub ShowMsg(byval num,byval title,byval msg,byval width)
  '// ���ؾ���/��ʾ��Ϣ
  '// stringhead��Html�ļ�ͷ��ȫ�ֱ���
  '---------------------------------------------------------------
  '// ��ʼ��
  if(width="")then width="100%"
  '---------------------------------------------------------------
  '// С�����е���ʾ��Ϣ
  if(num=0)then
     Response.Write(stringhead)
     Response.Write("<body bgcolor=#c6c6c6 leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>")
     Response.Write("<center><font class=black>"+msg+"</font></center>")
     Response.end
     end if
  '---------------------------------------------------------------
  '// �󴰿��е���ʾ��Ϣ��ǰ�벿��
  if(num=3)then Response.Write(Stringbody)
  '// �󴰿��е���ʾ��Ϣ��ǰ�벿��
  if(num=1 or num=3) then
     '����һ���Ǹռ���ȥ��
     'Response.Write("<center><table width='100%' border=0 cellspacing=0 cellpadding=0><tr><td>")
     Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.Write("  <tr align=center height=25>")
     Response.Write("    <td width=16 background='/image/002.gif'></td>")
     Response.Write("    <td background='/image/004.gif' align=left><img alt='���ں���' onclick='window.history.back();' style='{cursor : hand;}' border=0 src='/image/003.gif'></td>")
     Response.Write("    <td width=34 background='/image/005.gif'></td>")
     Response.Write("    <td width=" & len(title)*32 & " background='/image/006.gif'><font color=black ><b>"+title+"</b></td>")
     Response.Write("    <td width=44 background='/image/007.gif'></td>")
     Response.Write("    <td background='/image/004.gif' align=right><img align='absmiddle' alt='ˢ�´���' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/008.gif'></td>")
     Response.Write("    <td width=21 background='/image/010.gif'><img align='absmiddle' alt='�رմ���' onclick='history.back();' style='{cursor : hand;}' border='0' src='/image/009.gif'></td>")
     Response.Write("    </tr>")
     Response.Write("  </table>")
     Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     'Response.Write("<center><table width='100%' border=0 cellspacing=0 cellpadding=0>")
     Response.Write("  <tr bgcolor='#cfcfcf'>")
     Response.Write("    <td width=7 background='/image/011.gif'></td>")
     'Response.Write("    <td width='7' height='100%'><img src='/image/011.gif' style='height:100%;width:100%'></td>")
     Response.Write("    <td class=content height=100 valign=top>")
     Response.Write("<table width=100% border=0 cellspacing=5 cellpadding=5><tr><td>")
     end if
 
  '// �󴰿��е���ʾ��Ϣ���м䲿��
  if(num=3)then
     Response.Write(msg)
     end if
  '// �󴰿��е���ʾ��Ϣ�ĺ�벿��
  if(num=2 or num=3) then
     Response.Write("</td></tr></table>")
     Response.Write("    </td>")
     Response.Write("    <td width=7 background='/image/012.gif'></td></tr>")
     'Response.Write("    <td width='7' height='100%'><img src='/image/012.gif' style='height:100%;width:100%'></td></tr>")
     Response.Write("  </table>")
     Response.Write("<center><table width="+width+" border=0 cellpadding=0 cellspacing=0>")
     Response.Write("  <tr height=10>")
     Response.Write("    <td width=10 background='/image/013.gif'></td>")
     Response.Write("    <td background='/image/014.gif'><img src='/image/014.gif'></td>")
     Response.Write("    <td width=23 background='/image/015.gif'></td>")
     Response.Write("    </tr>")
     Response.Write("  </table>")
     '����һ���Ǹռ���ȥ��
     'Response.Write("  </td></tr></table>")
     end if
  
  '// �˳�
  if(num=3)then Response.end
  if(num=5 ) then
     '����һ���Ǹռ���ȥ��
     'Response.Write("<center><table width='100%' border=0 cellspacing=0 cellpadding=0><tr><td>")
     Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.Write("  <tr align=center height=25>")
     Response.Write("    <td width=16 background='/image/002.gif'></td>")
     Response.Write("    <td width=80 background='/image/004.gif' align=left><img alt='���ں���' onclick='window.history.back();' style='{cursor : hand;}' border=0 src='/image/003.gif'></td>")
     Response.Write("    <td width=" & len(title)*32 & " background='/image/006.gif'><font color=black ><b><div ID='draw' class='draw'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+title+"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div></b></td>")
     Response.Write("    <td background='/Ioffice/OA/Images/ShortMsg/smfra_top32.gif' align=right><img align='absmiddle' alt='ˢ�´���' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/008.gif'></td>")
     Response.Write("    <td width=21 background='/image/010.gif'></td>")
     Response.Write("    </tr>")
     Response.Write("  </table>")
     Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     'Response.Write("<center><table width='100%' border=0 cellspacing=0 cellpadding=0>")
     Response.Write("  <tr bgcolor='#cfcfcf'>")
     Response.Write("    <td width=7 background='/image/011.gif'></td>")
     'Response.Write("    <td width='7' height='100%'><img src='/image/011.gif' style='height:100%;width:100%'></td>")
     Response.Write("    <td class=content height=70 valign=top>")
     Response.Write("<table width=100% border=0 cellspacing=5 cellpadding=5 height=150 valign=top><tr><td valign=top>")
     end if
  if(num=4) then
     Response.Write(Stringbody3)
     Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.Write("  <tr align=center height=25>")
     Response.Write("    <td width=16 background='/image/002.gif'></td>")
     Response.Write("    <td width=100 background='/image/004.gif' align=left>&nbsp;</td>")
     Response.Write("    <td width=" & len(title)*32 & " background='/image/006.gif'><font color=black ><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+title+"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></td>")
     Response.Write("    <td background='/Ioffice/OA/Images/ShortMsg/smfra_top32.gif' align=right><img align='absmiddle' alt='ˢ�´���' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/008.gif'></td>")
     Response.Write("    <td width=21 background='/image/010.gif'></td>")
     Response.Write("    </tr>")
     Response.Write("  </table>")
     Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.Write("  <tr bgcolor='#cfcfcf'>")
     Response.Write("    <td width=7 background='/image/011.gif'></td>")
     Response.Write("    <td class=content height=50 valign=top>")
     Response.Write("<table width=100% border=0 cellspacing=5 cellpadding=5 height=150 valign=top><tr><td valign=top>")
     Response.Write(msg)
     Response.Write("</td></tr>")
     Response.Write("  </table>")
     Response.Write("<center><table width="+width+" border=0 cellpadding=0 cellspacing=0>")
     Response.Write("  <tr height=10>")
     Response.Write("    <td width=10 background='/image/013.gif'></td>")
     Response.Write("    <td background='/image/014.gif'><img src='/image/014.gif'></td>")
     Response.Write("    <td width=23 background='/image/015.gif'></td>")
     Response.Write("    </tr>")
     Response.Write("  </table>")
     Response.end
     end if
	'Ӧ����һЩ��ȳ�������Ĵ���,�ϰ벿��
	if(num=6) then
     '����һ���Ǹռ���ȥ��
     Response.Write("<center><table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td>")
     Response.Write("<center><table width='"+width+"' border='0' cellspacing='0' cellpadding='0'>")
     Response.Write("  <tr align='center' height='25'>")
     Response.Write("    <td width='16' background='/image/002.gif'></td>")
     Response.Write("    <td background='/image/004.gif' align='left'><img alt='���ں���' onclick='window.history.back();' style='{cursor : hand;}' border='0' src='/image/003.gif'></td>")
     Response.Write("    <td width='34' background='/image/005.gif'></td>")
     Response.Write("    <td width='" & len(title)*32 & "' background='/image/006.gif'><font color='black'><b>"+title+"</b></td>")
     Response.Write("    <td width='44' background='/image/007.gif'></td>")
     Response.Write("    <td background='/image/004.gif' align=right><img align='absmiddle' alt='ˢ�´���' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/008.gif'> <img align='absmiddle' alt='�رմ���' onclick='history.back();' style='{cursor : hand;}' border='0' src='/image/009.gif'></td>")
     Response.Write("    <td width='21' background='/image/010.gif'></td>")
     Response.Write("    </tr>")
     Response.Write("  </table>")
     'Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.Write("<center><table width='100%' border='0' cellspacing='0' cellpadding='0'>")
     Response.Write("  <tr bgcolor='white'>")
     'Response.Write("    <td width=7 background='/image/011.gif'></td>")
     Response.Write("    <td width='7' height='100%'><img src='/image/011.gif' style='height:100%;width:100%'></td>")
     Response.Write("    <td class='content' height='100' valign='top'>")
     Response.Write("<table width='100%' border='0' cellspacing='5' cellpadding='5'><tr><td>")
	end if     
	'Ӧ����һЩ��ȳ�������Ĵ���,�°벿��
	if(num=7) then
     Response.Write("</td></tr></table>")
     Response.Write("    </td>")
     'Response.Write("    <td width=7 background='/image/012.gif'></td></tr>")
     Response.Write("    <td width='7' height='100%'><img src='/image/012.gif' style='height:100%;width:100%'></td></tr>")
     Response.Write("  </table>")
     Response.Write("<center><table width='"+width+"' border='0' cellpadding='0' cellspacing='0'>")
     Response.Write("  <tr height='10'>")
     Response.Write("    <td width='10' background='/image/013.gif'></td>")
     Response.Write("    <td background='/image/014.gif'><img src='/image/014.gif'></td>")
     Response.Write("    <td width='23' background='/image/015.gif'></td>")
     Response.Write("    </tr>")
     Response.Write("  </table>")
     '����һ���Ǹռ���ȥ��
     Response.Write("  </td></tr></table>")	
	end if     
end sub

sub showmsg_old(byval num,byval title,byval msg,byval width)
  '// ���ؾ���/��ʾ��Ϣ
  '// stringhead��Html�ļ�ͷ��ȫ�ֱ���
  '---------------------------------------------------------------
  '// ��ʼ��
  if(width="")then width="100%"
  '---------------------------------------------------------------
  '// С�����е���ʾ��Ϣ
  if(num=0)then
     Response.write(stringhead)
     Response.write("<body bgcolor=#A0C0E8 leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>")
     Response.write("<center><font class=black>"+msg+"</font></center>")
     Response.end
     end if
  '---------------------------------------------------------------
  '// �󴰿��е���ʾ��Ϣ��ǰ�벿��
  if(num=3)then Response.write(Stringhead+Stringbody2)
  '// �󴰿��е���ʾ��Ϣ��ǰ�벿��
  if(num=1 or num=3)and(UserPZ4="��׼")then
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr align=center height=46>")
     Response.write("    <td width=16 background='/image/fra_top11.gif'></td>")
     Response.write("    <td background='/image/fra_top12.gif' align=left><img alt='���ں���' onclick='window.history.back();' style='{cursor : hand;}' border=0 src='/image/menu_back.gif'></td>")
     Response.write("    <td width=70 background='/image/fra_top13.gif'></td>")
     Response.write("    <td width=" & len(title)*32 & " background='/image/fra_top21.gif'><font color=white size=4><b>"+title+"</b></td>")
     Response.write("    <td width=60 background='/image/fra_top31.gif'></td>")
     Response.write("    <td background='/image/fra_top32.gif' align=right><img alt='ˢ�´���' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/menu_refresh.gif'></td>")
     Response.write("    <td width=16 background='/image/fra_top33.gif'></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr bgcolor=#d0d0d0>")
     Response.write("    <td width=9 background='/image/fra_mid1.gif'></td>")
     Response.write("    <td class=content height=100 valign=top>")
     Response.write("<table width=100% border=0 cellspacing=5 cellpadding=5><tr><td>")
     end if
  if(num=1 or num=3)and(UserPZ4="Ĭ��")then
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr align=center height=36>")
     Response.write("    <td width=17 background='/image/fra2_top11.gif'></td>")
     Response.write("    <td background='/image/fra2_top12.gif' align=left><img alt='���ں���' onclick='window.history.back();' style='{cursor : hand;}' border=0 src='/image/menu_back.gif'></td>")
     Response.write("    <td width=26 background='/image/fra2_top13.gif'></td>")
     Response.write("    <td width=" & len(title)*16 & " background='/image/fra2_top2.gif'><font color=white size=2><b>"+title+"</b></td>")
     Response.write("    <td width=23 background='/image/fra2_top31.gif'></td>")
     Response.write("    <td background='/image/fra2_top32.gif' align=right><img alt='ˢ�´���' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/menu_refresh.gif'></td>")
     Response.write("    <td width=22 background='/image/fra2_top33.gif'></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr>")
     Response.write("    <td width=17 background='/image/fra2_mid1.gif'></td>")
     Response.write("    <td class=content bgcolor=#d0d0d0 height=100 valign=top>")
     Response.write("<table width=100% bgcolor=#d0d0d0 border=0 cellspacing=0 cellpadding=1><tr><td>")
     end if
  '// �󴰿��е���ʾ��Ϣ���м䲿��
  if(num=3)then
     Response.write(msg)
     end if
  '// �󴰿��е���ʾ��Ϣ�ĺ�벿��
  if(num=2 or num=3)and(UserPZ4="��׼")then
     Response.write("</td></tr></table>")
     Response.write("    </td>")
     Response.write("    <td width=10 background='/image/fra_mid2.gif'></td></tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellpadding=0 cellspacing=0>")
     Response.write("  <tr height=10>")
     Response.write("    <td width=10 background='mage/fra_bot1.gif'></td>")
     Response.write("    <td background='/image/fra_bot2.gif'><img src='/image/fra_bot2.gif'></td>")
     Response.write("    <td width=10 background='/image/fra_bot3.gif'></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     end if
  if(num=2 or num=3)and(UserPZ4="Ĭ��")then
     Response.write("</td></tr></table>")
     Response.write("    </td>")
     Response.write("    <td width=22 background='/image/fra2_mid2.gif'></td></tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellpadding=0 cellspacing=0>")
     Response.write("  <tr height=18>")
     Response.write("    <td width=17 background='/image/fra2_bot1.gif'></td>")
     Response.write("    <td background='/image/fra2_bot2.gif'><img src='/image/fra2_bot2.gif'></td>")
     Response.write("    <td width=22 background='/image/fra2_bot3.gif'></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     end if
  '// �˳�
  if(num=3)then Response.end
end sub

sub showmsg1(byval num,byval title,byval msg,byval width)
  '// ���ؾ���/��ʾ��Ϣ
  '// stringhead��Html�ļ�ͷ��ȫ�ֱ���
  '---------------------------------------------------------------
  '// ��ʼ��
  if(width="")then width="100%"
  '---------------------------------------------------------------
  '// С�����е���ʾ��Ϣ
  if(num=0)then
     Response.write(stringhead)
     Response.write("<body bgcolor=#A0C0E8 leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>")
     Response.write("<center><font class=black>"+msg+"</font></center>")
     Response.end
     end if
  '---------------------------------------------------------------
  '// �󴰿��е���ʾ��Ϣ��ǰ�벿��
  if(num=3)then Response.write(Stringhead+Stringbody2)
  '// �󴰿��е���ʾ��Ϣ��ǰ�벿��
  if(num=1 or num=3)and(UserPZ4="��׼")then
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr align=center height=46>")
     Response.write("    <td width=16 ></td>")
     Response.write("    <td width=70 ></td>")
     Response.write("    <td width=" & len(title)*32 & " ><font color=black size=4><b>"+title+"</b></td>")
     Response.write("    <td width=60 ></td>")
     Response.write("    <td width=16 ></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr bgcolor=#FFFFFF>")
     Response.write("    <td width=9 ></td>")
     Response.write("    <td class=content height=100 valign=top>")
     Response.write("<table width=100% border=0 cellspacing=5 cellpadding=5><tr><td>")
     end if
  if(num=1 or num=3)and(UserPZ4="Ĭ��")then
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr align=center height=36>")
     Response.write("    <td width=17 ></td>")
     Response.write("    <td width=26 ></td>")
     Response.write("    <td width=" & len(title)*16 & " ><font color=black size=2><b>"+title+"</b></td>")
     Response.write("    <td width=23 ></td>")
     Response.write("    <td width=22 ></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr>")
     Response.write("    <td width=17 ></td>")
     Response.write("    <td class=content bgcolor=#FFFFFF height=100 valign=top>")
     Response.write("<table width=100% bgcolor=#FFFFFF border=0 cellspacing=0 cellpadding=1><tr><td>")
     end if
  '// �󴰿��е���ʾ��Ϣ���м䲿��
  if(num=3)then
     Response.write(msg)
     end if
  '// �󴰿��е���ʾ��Ϣ�ĺ�벿��
  if(num=2 or num=3)and(UserPZ4="��׼")then
     Response.write("</td></tr></table>")
     Response.write("    </td>")
     Response.write("    <td width=10 ></td></tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellpadding=0 cellspacing=0>")
     Response.write("  <tr height=10>")
     Response.write("    <td width=10 ></td>")
     Response.write("    <td ><img ></td>")
     Response.write("    <td width=10 ></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     end if
  if(num=2 or num=3)and(UserPZ4="Ĭ��")then
     Response.write("</td></tr></table>")
     Response.write("    </td>")
     Response.write("    <td width=22 ></td></tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellpadding=0 cellspacing=0>")
     Response.write("  <tr height=18>")
     Response.write("    <td width=17 ></td>")
     Response.write("    <td ></td>")
     Response.write("    <td width=22 ></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     end if
  '// �˳�
  if(num=3)then Response.end
end sub

function encode(byval pwd)
  '// ����任
  '---------------------------------------------------------------
  dim ii,leng,ch
  '---------------------------------------------------------------
  if(pwd="")then
     encode="PASSWORD"
     exit function
     end if
  if(pwd="PASSWORD")then
     encode=""
     exit function
     end if
  '---------------------------------------------------------------
  leng=len(pwd)
  for ii=1 to leng
      ch=mid(pwd,ii,1)
      '// if(asc(ch)>31 and asc(ch)<127)then pwd=left(pwd,ii-1)+chr(asc(ch) xor 15)+right(pwd,leng-ii)
      if(asc(ch)>0 and asc(ch)<128)then pwd=left(pwd,ii-1)+chr(asc(ch) xor 15)+right(pwd,leng-ii)
      next
  encode=pwd
  '---------------------------------------------------------------
end function
function intersect(byval ch,byval collect1,byval collect2)
  '// ����collect1,collect2�������ַ����Ľ���(chΪ���ϵķָ���)
  '---------------------------------------------------------------
  dim ii,temp
  '---------------------------------------------------------------
  if(collect1="" or collect2="")then
     intersect=""
     exit function
     end if
  '---------------------------------------------------------------
  temp=split(collect1,ch)
  collect2=ch+collect2+ch
  for ii=0 to ubound(temp)
      if(instr(collect2,ch+temp(ii)+ch)>0)then
         intersect=temp(ii)
         exit function
         end if
      next
  '---------------------------------------------------------------
end function
function newfilename(t,d,f)
  '// ���ݴ��롢��λ�����ŵȵó� �ʼ�������������ҳ�������û���Ƭ�ȵ��ļ����任����
  '// CoopDM�ǵ�λ�����ţ�ȫ�ֱ���
  '---------------------------------------------------------------
  if(USERMC="") then  CoopDM="00"
  '---------------------------------------------------------------
  if(t="mail")then newfilename=CoopDM+"m"+right("00000" & d,5)+"_"+f
  if(t="file")then newfilename=CoopDM+"f"+right("00000" & d,5)+"_"+f
  if(t="photo")then newfilename=CoopDM+"h"+right("00000" & d,5)+".jpg"
  '---------------------------------------------------------------
end function
function newfilelink(t,d,f)
  '---------------------------------------------------------------
  'newfilelink="action.asp?page=10&t="+t+"&d=" & d & "&f="+f
   newfilelink="file/"+newfilename(t,d,f)
  '---------------------------------------------------------------
end function
function SaveLog(CNMain,MC,msg)
  '// YHDM,YHMC���û�������û����ƣ�ȫ�ֱ���
  '---------------------------------------------------------------
  CNMain.Execute "insert into RZXX (RZSJ,RZYM,RZIP,RZMS) values('"+mytime()+"','"+MC+"','"+Request("REMOTE_ADDR")+"','"+msg+"')"
  '---------------------------------------------------------------
End Function
function getpurview(YHZL,YHQZ,QXGL,QXBJ,QXLL)
  '// ���� �û����࣬�û�Ⱥ���Ȩ�޲����ж��û��ķ���Ȩ��
  '---------------------------------------------------------------
  dim ii,temp
  '---------------------------------------------------------------
  getpurview=""
  if(YHZL="����Ա")then
     getpurview="������"
     exit function
  end if
  if(YHQZ="")then exit function
  temp=split(YHQZ,",")
  QXGL=","+QXGL+","
  QXBJ=","+QXBJ+","
  QXLL=","+QXLL+","
  for ii=0 to ubound(temp)
      if(instr(QXGL,","+temp(ii)+",")>0)then
         getpurview="������"
         exit function
         end if
      next
  for ii=0 to ubound(temp)
      if(instr(QXBJ,","+temp(ii)+",")>0)then
         getpurview="�༭��"
         exit function
         end if
      next
  for ii=0 to ubound(temp)
      if(instr(QXLL,","+temp(ii)+",")>0)then
         getpurview="�����"
         exit function
         end if
      next
  '---------------------------------------------------------------
end function
sub showorder(byval field,byval fname,byval title,byval link,byval order)
  '---------------------------------------------------------------
  dim dirmap
  '---------------------------------------------------------------
  dirmap="dir00.gif alt=��["+fname+"]��������"
  if(instr(order,field)>0)and(instr(order,"asc")>0)then dirmap="dir01.gif alt=��["+fname+"]��������"
  if(instr(order,field)>0)and(instr(order,"desc")>0)then dirmap="dir02.gif alt=��["+fname+"]��������"
  %><a href="<%=link%>&order=order by <%=field+" "%><%if(order<>"order by "+field+" desc")then%>desc<%else%>asc<%end if%>"><font color=black><%=title%><img src=image/<%=dirmap%> border=0 align=absmiddle></a><%
  '---------------------------------------------------------------
end sub
function isiden(iden)
  '// �жϱ�ʶ���Ƿ����Ҫ��
  '---------------------------------------------------------------
  dim ii,ch
  '---------------------------------------------------------------
  isiden=""
  if(iden="")then
     isiden="����Ϊ�ա�"
     exit function
     end if
  for ii=1 to len(iden)
      ch=mid(iden,ii,1)
      if(not((asc(ch)<=0)or(ch="_")or(ch>="0" and ch<="9")or(ch>="a" and ch<="z")or(ch>="A" and ch<="Z")))then
         isiden="ֻ�ܰ������֡���ĸ�����֡��»��ߡ��ո�"
         exit function
         end if
      next
  '---------------------------------------------------------------
end function
function fileupload(byval ftype,byval d,byref filename)
  '// ftype:"file","mail"
  '// ���أ��ɹ����ļ���
  '---------------------------------------------------------------
  '// �ÿؼ�iNotes.Upload�����ϴ�����
  Set oUpload = Server.CreateObject("iNotes.Upload")
  oUpload.FilePath=Server.MapPath("file")
  fileupload=false
  if(instr(oUpload.Request("file"),":\")>0)then
     filename=oUpload.FileName("file")
     fileupload=oUpload.SaveFile("file",newfilename(ftype,d,filename))
     end if
  '---------------------------------------------------------------
  '// �ϴ�����
  set oUpload = nothing
  '---------------------------------------------------------------
end function
function xinxiupload(byval ftype,byval d,byref filename)
  '// ftype:"file","mail"
  '// ���أ��ɹ����ļ���
  '---------------------------------------------------------------
  '// �ÿؼ�iNotes.Upload�����ϴ�����
  Set oUpload = Server.CreateObject("iNotes.Upload")
  oUpload.FilePath=Server.MapPath("xinxiup")
  xinxiupload=false
  if(instr(oUpload.Request("xinxiup"),":\")>0)then
     filename=oUpload.FileName("xinxiup")
     xinxiupload=oUpload.SaveFile("xinxiup",newfilename(ftype,d,filename))
     end if
  '---------------------------------------------------------------
  '// �ϴ�����
  set oUpload = nothing
  '---------------------------------------------------------------
end function
sub checkpurview(byval z,byval QXZL,byref purview)
  '// �ж�ͨѶ¼������������������������ҳ���ղؼе�Ȩ��
  '---------------------------------------------------------------
  dim RSMain
  Set RSMain =Server.Createobject("ADODB.Recordset")
  '---------------------------------------------------------------
  if(z=UserMC)then purview="������"
  if(z="[���]")then  '// ������ĵ�
     purview="�༭��"
     end if
  if(z="[�ڲ�]")then  '// �ڲ�ͨѶ¼
     purview="�����"
     if(UserZL="����Ա")then purview="������"
     end if
  if(z="" or instr(z,chr(5))>0)then purview="�����"
  if(z<>UserMC)and(z<>"[�ڲ�]")and(z<>"[���]")and(z<>"")and(instr(z,chr(5))<=0)then  '// �Զ����
     RSMain.Open "select * from QXXX where (QXZL='"+QXZL+"')and(QXMC='"+z+"')",Connect,1,3 '// �������޸ġ���ɾ��
     if(RSMain.eof)then call showmsg(3,"������Ϣ","�벻Ҫ�÷������ķ����򿪲�����Ŀ¼��","")
     purview=getpurview(UserZL,UserP2,RSMain("QXGL"),RSMain("QXBJ"),RSMain("QXLL"))
     if(purview="")then call showmsg(3,"������Ϣ","�벻Ҫ�÷������ķ�������û�з���Ȩ�޵�Ŀ¼��","")
     RSMain.Close
     end if
  '---------------------------------------------------------------
end sub
function convertmsg(byval msg,byval para)
  '---------------------------------------------------------------
  dim temp
  '---------------------------------------------------------------
  temp=ucase(msg)
  if(instr(temp,"<BR>")<=0)and(instr(temp,"<P>")<=0)and(instr(temp,"<TABLE")<=0)and(instr(temp,"<HTML>")<=0)then
     msg=replace(msg,"  ","��")
     msg=replace(msg,chr(13)+chr(10),"<BR>")
     end if
  convertmsg=msg
  '---------------------------------------------------------------
end function

function Mailto(RSMain,YHMC,MailSX,MailBT,MailNR,MailURL)
  '// ���ͼ��ż�
  '---------------------------------------------------------------
  '// д�ż�
  'ConnectOffice = "Provider=SQLOLEDB;Server=RH28XX60;DataBase=ioffice;Uid=sa;Pwd=weboaadmin2003;"
  'Set CNMainOFFICE=Server.CreateObject("ADODB.Connection")
  'CNMainOFFICE.open ConnectOffice,"",""
  RSMain.Open "select * from MailXX ",Connect,1,3 '// �������޸ġ���ɾ��
  RSMain.Addnew
  RSMain("MailSJ")=mytime()
  RSMain("MailJB")="��ͨ"
  RSMain("MailFX")=YHMC
  RSMain("MailSX")=MailSX
  RSMain("MailBT")=MailBT
  RSMain("MailNR")=MailNR
  RsMain("MailBM")=UserBM
  RsMain("MailURL")=MailURL
  RsMain("MailBZ")="WEBOA֪ͨ"
  RSMain.Update
  MailDM=RSMain("MailDM")
  RSMain.Close
  '---------------------------------------------------------------
  '// д�����˼�¼
  if(YHMC<>"")then
     'CNMainOffice.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFLL) values('����'," & MailDM & ",'"+YHMC+"','������','������',1)")
     CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFLL,SFBM,SFBZ) values('����'," & MailDM & ",'"+YHMC+"','��֪ͨ��','��֪ͨ��',1,'"+UserBM+"','WEBOA֪ͨ')")
  end if
  '---------------------------------------------------------------
  '// д�ʼ��������˼�¼
  temp="'"+replace(MailSX,",","','")+"'"
  'RSMain.open "select cstr(YHMC) from YHXX where YHMC in ("+temp+") union select cstr(QZYM) from QZXX where QZMC in("+temp+")",Connect,1,1 '// ��ȡ
  RSMain.open "select YHMC from YHXX where YHMC in ("+temp+") union select QZYM from QZXX where QZMC in("+temp+")",Connect,1,1 '// ��ȡ
  temp="*"
  while (not RSMain.Eof)
     temp=temp+","+RSMain(0)
     RSMain.Movenext
     wend
  RSMain.Close
  temp="'"+replace(temp,",","','")+"'"
  CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFBZ) select '����'," & MailDM & ",YHMC,'��֪ͨ��','��֪ͨ��','WEBOA֪ͨ' from YHXX where (YHMC in ("+temp+"))and(YHMC<>'"+YHMC+"')")
  'CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFBZ) select '����'," & MailDM & ",YHMC,'��֪ͨ��','��֪ͨ��' from YHXX where (YHMC in ("+temp+"))and(YHMC<>'"+YHMC+"')")
  
  '---------------------------------------------------------------
end function

function Mailto_old(RSMain,YHMC,MailSX,MailBT,MailNR)
  '// ���ͼ��ż�
  '---------------------------------------------------------------
  '// д�ż�
  RSMain.Open "select * from MailXX ",Connect,1,3 '// �������޸ġ���ɾ��
  RSMain.Addnew
  RSMain("MailSJ")=mytime()
  RSMain("MailJB")="��ͨ"
  RSMain("MailFX")=YHMC
  RSMain("MailSX")=MailSX
  RSMain("MailBT")=MailBT
  RSMain("MailNR")=MailNR
  RSMain.Update
  MailDM=RSMain("MailDM")
  RSMain.Close
  '---------------------------------------------------------------
  '// д�����˼�¼
  if(YHMC<>"")then
     CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFLL) values('����'," & MailDM & ",'"+YHMC+"','������','������',1)")
     end if
  '---------------------------------------------------------------
  '// д�ʼ��������˼�¼
  temp="'"+replace(MailSX,",","','")+"'"
  'RSMain.open "select cstr(YHMC) from YHXX where YHMC in ("+temp+") union select cstr(QZYM) from QZXX where QZMC in("+temp+")",Connect,1,1 '// ��ȡ
  RSMain.open "select YHMC from YHXX where YHMC in ("+temp+") union select QZYM from QZXX where QZMC in("+temp+")",Connect,1,1 '// ��ȡ
  temp="*"
  while (not RSMain.Eof)
     temp=temp+","+RSMain(0)
     RSMain.Movenext
     wend
  RSMain.Close
  temp="'"+replace(temp,",","','")+"'"
  CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2) select '����'," & MailDM & ",YHMC,'�ռ���','�ռ���' from YHXX where (YHMC in ("+temp+"))and(YHMC<>'"+YHMC+"')")
  '---------------------------------------------------------------
end function

  '---------------------------------------------------------------
  CoopXX=session("sys.CoopXX")
  CoopDM=fieldget(CoopXX,1,chr(5))
  CoopMC=fieldget(CoopXX,2,chr(5))
  Connect=fieldget(CoopXX,3,chr(5))
  '---------------------------------------------------------------
  UserXX=session("sys.UserXX")
  UserDM=fieldget(UserXX,1,chr(5))  '// ����
  UserDL=fieldget(UserXX,2,chr(5))  '// ��¼��
  UserMC=fieldget(UserXX,3,chr(5))  '// �û���
  UserBM=fieldget(UserXX,4,chr(5))  '// ����
  UserZL=fieldget(UserXX,5,chr(5))  '// ����
  UserP0=session("sys.UserP0")  '// �ҵ�����
  UserP1=session("sys.UserP1") '// Ⱥ��
  UserP2=session("sys.UserP2") '// Ⱥ��
  UserPZ=session("sys.UserPZ")  '// ����ѡ��
  UserPZ4=fieldget(UserPZ,4,chr(5))
  UserZJ=session("sys.UserZJ")
  UserBMCODE=session("sys.UserBMCODE")
    
  IF UserMC<>"" and UserBM<>"" THEN
    set bmco=server.createobject("ADODB.Recordset")
    bmco.open "select * from bmxx where qzmc='"&userbm&"'",connect,1,3
      if not bmco.eof then 
         bmcode=bmco("bmcode")
         UserBMCODE=bmco("bmcode")
         userqzcode=bmco("qzcode")
      end if  
    bmco.close:set bmco=nothing
  
    set YHZJ=server.createobject("ADODB.Recordset")
    YHZJ.OPEN "select * from ryjbqk where a0101='"&usermc&"' and za0101='"&bmcode&"'",connect,1,3
        if not YHZJ.eof then UserZJ=YHZJ("A0221") 
    YHZJ.close
    set YHZJ=nothing
  END IF 
  
  if(UserPZ4<>"��׼")then UserPZ4="Ĭ��"
  thisfile=Request.ServerVariables("SCRIPT_NAME")
  '---------------------------------------------------------------
  '// StrinfDetect="<SCRIPT language=JavaScript><!--"+chr(13)+chr(10)+"if(parent.logo.window.oa.alt!='��ҵӦ��ƽ̨') window.location.href='login.asp?page=02';"+chr(13)+chr(10)+"//--></SCRIPT>"+chr(13)+chr(10)
  StringAlarm="�Ͻ���δ��¼��ʱ״����ʹ����Ϣϵͳ��<br>Ҫ���µ�¼���<a target=_parent href='login.asp'>����</a>��"
  StringWind="<center><table border=0 width=100% cellspacing=0 cellpadding=1 bgcolor=#FFFFFF><tr><td><nolayer><iframe name=_edit height=15 width=100% marginwidth=0 marginheight=0 scrolling=no frameborder=0></iframe></nolayer></td></tr></table></center>"
  StringHead="<HTML><HEAD><TITLE>"+CoopMC+"Iasi Information System V2.0</TITLE><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><LINK rel=""stylesheet"" href=""style.css""></HEAD>"
  StringBody="<body bgcolor=#ffffff leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>"
  StringBody2="<body bgcolor=#ffffff leftmargin=5 rightmargin=5 topmargin=5 bottommargin=5>"
  if(UserPZ4="��׼")then
     StringWind="<center><table border=1 width=100% cellspacing=0 cellpadding=1 bgcolor=#A0C0E8><tr><td><nolayer><iframe name=_edit height=15 width=100% marginwidth=0 marginheight=0 scrolling=no frameborder=0></iframe></nolayer></td></tr></table></center>"
     StringBody="<body _bgcolor=#FFF9E6 bgcolor=#4488CC leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>"
     StringBody2="<body _bgcolor=#FFF9E6 bgcolor=#4488CC leftmargin=5 rightmargin=5 topmargin=5 bottommargin=5>"
     end if
  '---------------------------------------------------------------
%>
