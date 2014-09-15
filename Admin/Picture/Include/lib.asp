<%
'---------------------------------------------------------------
  page=Request("page")  
'Connect2="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=NEWSCBRCGD;Data Source=oa-server2;Pwd=weboaadmin2004"
Connect2="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=PBC;Data Source=192.168.0.201;Pwd="
  Set CNMain2=Server.CreateObject("ADODB.Connection")
  CNMain2.open Connect2,"",""
' ---------------------------------------------------------------

' 在Domino服务器上验证用户是否合法
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
        Response.Write "无法验证！此地址不可访问：" + DOMINO_LOGIN_URL + "<br><br>"
    end if
    'Response.Write Text
end function

function FieldGet(byval mystr,byval Pos,byval delimiter)
    '---------------------------------------------------------------
    '// 定义局部变量
    Dim count,leng,index,pindex
    '---------------------------------------------------------------
    '// 若输入字符串为空,或查找域号小于1,则返回空字符串
    If (mystr = "") Or (delimiter = "") Or (Pos < 1) Then
       FieldGet = ""
       Exit Function
       End If
    '---------------------------------------------------------------
    '// 循环查找,将指针指向要读取域的起始位置
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
    '//定义局部变量
    Dim count,leng,index
    '---------------------------------------------------------------
    '// 若输入字符串为空,则返回域数目为1
    If (mystr= "") Or (delimiter ="") Then
       FieldCount = 1
       Exit Function
       End If
    '---------------------------------------------------------------
    '// 循环查找,将指针指向要读取域的起始位置
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
    '//按date1->date2->date3->date4顺序，选择有值得数
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
  '// 按指定的格式显示日期
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
  '// 按指定的格式显示日期/时间
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
  '// 按指定的格式显示日期/时间
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
  '// 显示yyyymmdd日期格式
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
  '// 根据参数获得指定的SQL查询语句
  '// 如果用户输入的是要查询的[关键字]，则得到相应的查询语句
  '---------------------------------------------------------------
  '// table:表名
  '// key  :需要检索的关键字
  '// flist:需要检索关键字的字段
  '// cond2:附加查询条件
  '---------------------------------------------------------------
  dim ii,cond,temp,temp2
  '---------------------------------------------------------------
  '// 如果关键字为空,则根据返回查询语句
  if(key="")then
     cond="select * from "+table+" "+order
     if(cond2<>"")then cond="select * from "+table+" where ("+cond2+") "+order
     GetCommand=cond
     exit function
     end if
  '---------------------------------------------------------------
  '// 如果用户输入的是[查询条件]，则得到相应的查询语句
  temp2=left(lcase(trim(key)),6)
  '// 根据where条件得到相应的查询语句
  if(temp2="where ")then
     temp2=replace(trim(key),"""","'")
     cond="select * from "+table+" where ("+right(temp2,len(temp2)-6)+")"
     if(cond2<>"")then cond=cond+"and("+cond2+")"
     GetCommand=cond+" "+order
     exit function
     end if
  '// 根据order条件得到相应的查询语句
  if(temp2="order ")then
     temp2=trim(key)
     cond="select * from "+table+" order by "+right(temp2,len(temp2)-9)
     if(cond2<>"")then cond="select * from "+table+" where ("+cond2+") "+key
     GetCommand=cond
     exit function
     end if
  '// 根据top条件得到相应的查询语句
  if(left(temp2,4)="top ")then
     cond="select "+key+" * from "+table
     if(cond2<>"")then cond="select "+key+" * from "+table+" where ("+cond2+")"
     GetCommand=cond+" "+order
     exit function
     end if
  '---------------------------------------------------------------
  '// 如果关键字不为空,则组合需要检索的字段序列
  flist=replace(flist,",","+'~'+")
  '---------------------------------------------------------------
  '// 得到单关键字的查询条件
  cond=" where (charindex('"+key+"',"+flist+")>0)"
  '// 得到[或]条件的多关键字检索语句
  if(instr(key,"|")>0)then
     temp=split(key,"|")
     cond="(charindex('"+temp(0)+"',"+flist+")>0)"
     for ii=1 to ubound(temp)
         cond=cond+"or(charindex('"+temp(ii)+"',"+flist+")>0)"
         next
     cond=" where ("+cond+")"
     end if
  '// 得到[并]条件的多关键字检索语句
  if(instr(key,"/")>0)then
     temp=split(key,"/")
     cond="(charindex('"+temp(0)+"',"+flist+")>0)"
     for ii=1 to ubound(temp)
         cond=cond+"and(charindex('"+temp(ii)+"',"+flist+")>0)"
         next
     cond=" where ("+cond+")"
     end if
  '---------------------------------------------------------------
  '// 根据cond条件合成完整的查询语句
  GetCommand="select * from "+table+cond+" "+order
  if(cond2<>"")then GetCommand="select * from "+table+cond+" and ("+cond2+") "+order
  '---------------------------------------------------------------
end function
sub Setbar(link,beginstring,PageIndex,PageCount,PageSize,RecordCount,endstring)
  '// 显示[上一页][下一页]的导航条
  '---------------------------------------------------------------
  bgcolor1="bgcolor=#FFFFFF" : bgcolor2="bgcolor=#000000"
  if(UserPZ4<>"标准")then
     bgcolor1=" bgcolor=#546E86"
     bgcolor2=" bgcolor=#546E86"
     end if
  Response.Write("<center><table width=100% cellspacing=0 cellpadding=1 "+bgcolor1+"><tr height=20 "+bgcolor2+"><td>")
  'Response.Write("<center><table border=1 width=100% cellspacing=0 cellpadding=1 "+bgcolor1+"><tr valign=middle height=22"+bgcolor2+"><td>")
  Response.Write(beginstring)
  Response.Write("<font color=#ffffff>[第"& PageIndex & "/" & PageCount & "页]")
  'Response.Write("<font color=#ffffff>[每页" & PageSize & "条/共" & RecordCount & "条记录]")
  if(PageIndex<=1)then Response.Write("<font color=#ffffff>[首页][上一页]")
  if(PageIndex>1)then Response.Write("<a href="""+link+"&n=1""><font color=#ffffff>[首页]</a><a href="""+link+"&n=" & PageIndex-1 & """><font color=#ffffff>[上一页]</a>")
  if(PageIndex>=PageCount)then Response.Write("<font color=#ffffff>[下一页][尾页]")
  if(PageIndex<PageCount)then Response.Write("<a href="""+link+"&n=" & (PageIndex+1) & """><font color=#ffffff>[下一页]</a><a href="""+link+"&n=" & PageCount & """><font color=#ffffff>[尾页]</a>")
  Response.Write(endstring)
  '---------------------------------------------------------------
end sub
sub alarm(byval msg)
  '// 返回警告/提示信息
  '// stringhead是Html文件头，全局变量
  '---------------------------------------------------------------
  bgcolor1=" bgcolor=#A0C0E8"
  if(UserPZ4<>"标准")then bgcolor1=" bgcolor=#FFFFE8"
  Response.write(stringhead)
  Response.write("<body "+bgcolor1+" leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>")
  Response.write("<center><font class=black>"+msg+"</font></center>")
  if(fieldget(session("sys.UserPZ"),8,chr(5))="是")then
     Response.write("<SCRIPT language=JavaScript><!--"+chr(13)+chr(10)+"alert('"+msg+"');"+chr(13)+chr(10)+"//--></SCRIPT>"+chr(13)+chr(10))
  end if
  Response.end
  '---------------------------------------------------------------
end sub

sub ShowMsg(byval num,byval title,byval msg,byval width)
  '// 返回警告/提示信息
  '// stringhead是Html文件头，全局变量
  '---------------------------------------------------------------
  '// 初始化
  if(width="")then width="100%"
  '---------------------------------------------------------------
  '// 小窗口中的提示信息
  if(num=0)then
     Response.Write(stringhead)
     Response.Write("<body bgcolor=#c6c6c6 leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>")
     Response.Write("<center><font class=black>"+msg+"</font></center>")
     Response.end
     end if
  '---------------------------------------------------------------
  '// 大窗口中的提示信息的前半部分
  if(num=3)then Response.Write(Stringbody)
  '// 大窗口中的提示信息的前半部分
  if(num=1 or num=3) then
     '下面一行是刚加上去的
     'Response.Write("<center><table width='100%' border=0 cellspacing=0 cellpadding=0><tr><td>")
     Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.Write("  <tr align=center height=25>")
     Response.Write("    <td width=16 background='/image/002.gif'></td>")
     Response.Write("    <td background='/image/004.gif' align=left><img alt='窗口后退' onclick='window.history.back();' style='{cursor : hand;}' border=0 src='/image/003.gif'></td>")
     Response.Write("    <td width=34 background='/image/005.gif'></td>")
     Response.Write("    <td width=" & len(title)*32 & " background='/image/006.gif'><font color=black ><b>"+title+"</b></td>")
     Response.Write("    <td width=44 background='/image/007.gif'></td>")
     Response.Write("    <td background='/image/004.gif' align=right><img align='absmiddle' alt='刷新窗口' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/008.gif'></td>")
     Response.Write("    <td width=21 background='/image/010.gif'><img align='absmiddle' alt='关闭窗口' onclick='history.back();' style='{cursor : hand;}' border='0' src='/image/009.gif'></td>")
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
 
  '// 大窗口中的提示信息的中间部分
  if(num=3)then
     Response.Write(msg)
     end if
  '// 大窗口中的提示信息的后半部分
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
     '下面一行是刚加上去的
     'Response.Write("  </td></tr></table>")
     end if
  
  '// 退出
  if(num=3)then Response.end
  if(num=5 ) then
     '下面一行是刚加上去的
     'Response.Write("<center><table width='100%' border=0 cellspacing=0 cellpadding=0><tr><td>")
     Response.Write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.Write("  <tr align=center height=25>")
     Response.Write("    <td width=16 background='/image/002.gif'></td>")
     Response.Write("    <td width=80 background='/image/004.gif' align=left><img alt='窗口后退' onclick='window.history.back();' style='{cursor : hand;}' border=0 src='/image/003.gif'></td>")
     Response.Write("    <td width=" & len(title)*32 & " background='/image/006.gif'><font color=black ><b><div ID='draw' class='draw'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+title+"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div></b></td>")
     Response.Write("    <td background='/Ioffice/OA/Images/ShortMsg/smfra_top32.gif' align=right><img align='absmiddle' alt='刷新窗口' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/008.gif'></td>")
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
     Response.Write("    <td background='/Ioffice/OA/Images/ShortMsg/smfra_top32.gif' align=right><img align='absmiddle' alt='刷新窗口' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/008.gif'></td>")
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
	'应用于一些宽度超出常规的窗口,上半部分
	if(num=6) then
     '下面一行是刚加上去的
     Response.Write("<center><table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td>")
     Response.Write("<center><table width='"+width+"' border='0' cellspacing='0' cellpadding='0'>")
     Response.Write("  <tr align='center' height='25'>")
     Response.Write("    <td width='16' background='/image/002.gif'></td>")
     Response.Write("    <td background='/image/004.gif' align='left'><img alt='窗口后退' onclick='window.history.back();' style='{cursor : hand;}' border='0' src='/image/003.gif'></td>")
     Response.Write("    <td width='34' background='/image/005.gif'></td>")
     Response.Write("    <td width='" & len(title)*32 & "' background='/image/006.gif'><font color='black'><b>"+title+"</b></td>")
     Response.Write("    <td width='44' background='/image/007.gif'></td>")
     Response.Write("    <td background='/image/004.gif' align=right><img align='absmiddle' alt='刷新窗口' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/008.gif'> <img align='absmiddle' alt='关闭窗口' onclick='history.back();' style='{cursor : hand;}' border='0' src='/image/009.gif'></td>")
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
	'应用于一些宽度超出常规的窗口,下半部分
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
     '下面一行是刚加上去的
     Response.Write("  </td></tr></table>")	
	end if     
end sub

sub showmsg_old(byval num,byval title,byval msg,byval width)
  '// 返回警告/提示信息
  '// stringhead是Html文件头，全局变量
  '---------------------------------------------------------------
  '// 初始化
  if(width="")then width="100%"
  '---------------------------------------------------------------
  '// 小窗口中的提示信息
  if(num=0)then
     Response.write(stringhead)
     Response.write("<body bgcolor=#A0C0E8 leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>")
     Response.write("<center><font class=black>"+msg+"</font></center>")
     Response.end
     end if
  '---------------------------------------------------------------
  '// 大窗口中的提示信息的前半部分
  if(num=3)then Response.write(Stringhead+Stringbody2)
  '// 大窗口中的提示信息的前半部分
  if(num=1 or num=3)and(UserPZ4="标准")then
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr align=center height=46>")
     Response.write("    <td width=16 background='/image/fra_top11.gif'></td>")
     Response.write("    <td background='/image/fra_top12.gif' align=left><img alt='窗口后退' onclick='window.history.back();' style='{cursor : hand;}' border=0 src='/image/menu_back.gif'></td>")
     Response.write("    <td width=70 background='/image/fra_top13.gif'></td>")
     Response.write("    <td width=" & len(title)*32 & " background='/image/fra_top21.gif'><font color=white size=4><b>"+title+"</b></td>")
     Response.write("    <td width=60 background='/image/fra_top31.gif'></td>")
     Response.write("    <td background='/image/fra_top32.gif' align=right><img alt='刷新窗口' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/menu_refresh.gif'></td>")
     Response.write("    <td width=16 background='/image/fra_top33.gif'></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr bgcolor=#d0d0d0>")
     Response.write("    <td width=9 background='/image/fra_mid1.gif'></td>")
     Response.write("    <td class=content height=100 valign=top>")
     Response.write("<table width=100% border=0 cellspacing=5 cellpadding=5><tr><td>")
     end if
  if(num=1 or num=3)and(UserPZ4="默认")then
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr align=center height=36>")
     Response.write("    <td width=17 background='/image/fra2_top11.gif'></td>")
     Response.write("    <td background='/image/fra2_top12.gif' align=left><img alt='窗口后退' onclick='window.history.back();' style='{cursor : hand;}' border=0 src='/image/menu_back.gif'></td>")
     Response.write("    <td width=26 background='/image/fra2_top13.gif'></td>")
     Response.write("    <td width=" & len(title)*16 & " background='/image/fra2_top2.gif'><font color=white size=2><b>"+title+"</b></td>")
     Response.write("    <td width=23 background='/image/fra2_top31.gif'></td>")
     Response.write("    <td background='/image/fra2_top32.gif' align=right><img alt='刷新窗口' onclick='location.reload();' style='{cursor : hand;}' border=0 src='/image/menu_refresh.gif'></td>")
     Response.write("    <td width=22 background='/image/fra2_top33.gif'></td>")
     Response.write("    </tr>")
     Response.write("  </table>")
     Response.write("<center><table width="+width+" border=0 cellspacing=0 cellpadding=0>")
     Response.write("  <tr>")
     Response.write("    <td width=17 background='/image/fra2_mid1.gif'></td>")
     Response.write("    <td class=content bgcolor=#d0d0d0 height=100 valign=top>")
     Response.write("<table width=100% bgcolor=#d0d0d0 border=0 cellspacing=0 cellpadding=1><tr><td>")
     end if
  '// 大窗口中的提示信息的中间部分
  if(num=3)then
     Response.write(msg)
     end if
  '// 大窗口中的提示信息的后半部分
  if(num=2 or num=3)and(UserPZ4="标准")then
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
  if(num=2 or num=3)and(UserPZ4="默认")then
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
  '// 退出
  if(num=3)then Response.end
end sub

sub showmsg1(byval num,byval title,byval msg,byval width)
  '// 返回警告/提示信息
  '// stringhead是Html文件头，全局变量
  '---------------------------------------------------------------
  '// 初始化
  if(width="")then width="100%"
  '---------------------------------------------------------------
  '// 小窗口中的提示信息
  if(num=0)then
     Response.write(stringhead)
     Response.write("<body bgcolor=#A0C0E8 leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>")
     Response.write("<center><font class=black>"+msg+"</font></center>")
     Response.end
     end if
  '---------------------------------------------------------------
  '// 大窗口中的提示信息的前半部分
  if(num=3)then Response.write(Stringhead+Stringbody2)
  '// 大窗口中的提示信息的前半部分
  if(num=1 or num=3)and(UserPZ4="标准")then
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
  if(num=1 or num=3)and(UserPZ4="默认")then
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
  '// 大窗口中的提示信息的中间部分
  if(num=3)then
     Response.write(msg)
     end if
  '// 大窗口中的提示信息的后半部分
  if(num=2 or num=3)and(UserPZ4="标准")then
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
  if(num=2 or num=3)and(UserPZ4="默认")then
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
  '// 退出
  if(num=3)then Response.end
end sub

function encode(byval pwd)
  '// 口令变换
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
  '// 计算collect1,collect2这两个字符串的交集(ch为集合的分隔符)
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
  '// 根据代码、单位索引号等得出 邮件附件、部门网页附件、用户照片等的文件名变换规则
  '// CoopDM是单位索引号，全局变量
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
  '// YHDM,YHMC是用户代码和用户名称，全局变量
  '---------------------------------------------------------------
  CNMain.Execute "insert into RZXX (RZSJ,RZYM,RZIP,RZMS) values('"+mytime()+"','"+MC+"','"+Request("REMOTE_ADDR")+"','"+msg+"')"
  '---------------------------------------------------------------
End Function
function getpurview(YHZL,YHQZ,QXGL,QXBJ,QXLL)
  '// 根据 用户种类，用户群组和权限参数判断用户的访问权限
  '---------------------------------------------------------------
  dim ii,temp
  '---------------------------------------------------------------
  getpurview=""
  if(YHZL="管理员")then
     getpurview="管理者"
     exit function
  end if
  if(YHQZ="")then exit function
  temp=split(YHQZ,",")
  QXGL=","+QXGL+","
  QXBJ=","+QXBJ+","
  QXLL=","+QXLL+","
  for ii=0 to ubound(temp)
      if(instr(QXGL,","+temp(ii)+",")>0)then
         getpurview="管理者"
         exit function
         end if
      next
  for ii=0 to ubound(temp)
      if(instr(QXBJ,","+temp(ii)+",")>0)then
         getpurview="编辑者"
         exit function
         end if
      next
  for ii=0 to ubound(temp)
      if(instr(QXLL,","+temp(ii)+",")>0)then
         getpurview="浏览者"
         exit function
         end if
      next
  '---------------------------------------------------------------
end function
sub showorder(byval field,byval fname,byval title,byval link,byval order)
  '---------------------------------------------------------------
  dim dirmap
  '---------------------------------------------------------------
  dirmap="dir00.gif alt=按["+fname+"]降序排列"
  if(instr(order,field)>0)and(instr(order,"asc")>0)then dirmap="dir01.gif alt=按["+fname+"]降序排列"
  if(instr(order,field)>0)and(instr(order,"desc")>0)then dirmap="dir02.gif alt=按["+fname+"]升序排列"
  %><a href="<%=link%>&order=order by <%=field+" "%><%if(order<>"order by "+field+" desc")then%>desc<%else%>asc<%end if%>"><font color=black><%=title%><img src=image/<%=dirmap%> border=0 align=absmiddle></a><%
  '---------------------------------------------------------------
end sub
function isiden(iden)
  '// 判断标识符是否符合要求
  '---------------------------------------------------------------
  dim ii,ch
  '---------------------------------------------------------------
  isiden=""
  if(iden="")then
     isiden="不能为空。"
     exit function
     end if
  for ii=1 to len(iden)
      ch=mid(iden,ii,1)
      if(not((asc(ch)<=0)or(ch="_")or(ch>="0" and ch<="9")or(ch>="a" and ch<="z")or(ch>="A" and ch<="Z")))then
         isiden="只能包含汉字、字母、数字、下划线、空格。"
         exit function
         end if
      next
  '---------------------------------------------------------------
end function
function fileupload(byval ftype,byval d,byref filename)
  '// ftype:"file","mail"
  '// 返回：成功的文件名
  '---------------------------------------------------------------
  '// 用控件iNotes.Upload功能上传数据
  Set oUpload = Server.CreateObject("iNotes.Upload")
  oUpload.FilePath=Server.MapPath("file")
  fileupload=false
  if(instr(oUpload.Request("file"),":\")>0)then
     filename=oUpload.FileName("file")
     fileupload=oUpload.SaveFile("file",newfilename(ftype,d,filename))
     end if
  '---------------------------------------------------------------
  '// 上传结束
  set oUpload = nothing
  '---------------------------------------------------------------
end function
function xinxiupload(byval ftype,byval d,byref filename)
  '// ftype:"file","mail"
  '// 返回：成功的文件名
  '---------------------------------------------------------------
  '// 用控件iNotes.Upload功能上传数据
  Set oUpload = Server.CreateObject("iNotes.Upload")
  oUpload.FilePath=Server.MapPath("xinxiup")
  xinxiupload=false
  if(instr(oUpload.Request("xinxiup"),":\")>0)then
     filename=oUpload.FileName("xinxiup")
     xinxiupload=oUpload.SaveFile("xinxiup",newfilename(ftype,d,filename))
     end if
  '---------------------------------------------------------------
  '// 上传结束
  set oUpload = nothing
  '---------------------------------------------------------------
end function
sub checkpurview(byval z,byval QXZL,byref purview)
  '// 判断通讯录，公告栏，讨论区，部门网页，收藏夹的权限
  '---------------------------------------------------------------
  dim RSMain
  Set RSMain =Server.Createobject("ADODB.Recordset")
  '---------------------------------------------------------------
  if(z=UserMC)then purview="管理者"
  if(z="[借出]")then  '// 借出的文档
     purview="编辑者"
     end if
  if(z="[内部]")then  '// 内部通讯录
     purview="浏览者"
     if(UserZL="管理员")then purview="管理者"
     end if
  if(z="" or instr(z,chr(5))>0)then purview="浏览者"
  if(z<>UserMC)and(z<>"[内部]")and(z<>"[借出]")and(z<>"")and(instr(z,chr(5))<=0)then  '// 自定义的
     RSMain.Open "select * from QXXX where (QXZL='"+QXZL+"')and(QXMC='"+z+"')",Connect,1,3 '// 新增、修改、或删除
     if(RSMain.eof)then call showmsg(3,"警告信息","请不要用非正常的方法打开不存在目录。","")
     purview=getpurview(UserZL,UserP2,RSMain("QXGL"),RSMain("QXBJ"),RSMain("QXLL"))
     if(purview="")then call showmsg(3,"警告信息","请不要用非正常的方法打开您没有访问权限的目录。","")
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
     msg=replace(msg,"  ","　")
     msg=replace(msg,chr(13)+chr(10),"<BR>")
     end if
  convertmsg=msg
  '---------------------------------------------------------------
end function

function Mailto(RSMain,YHMC,MailSX,MailBT,MailNR,MailURL)
  '// 发送简单信件
  '---------------------------------------------------------------
  '// 写信件
  'ConnectOffice = "Provider=SQLOLEDB;Server=RH28XX60;DataBase=ioffice;Uid=sa;Pwd=weboaadmin2003;"
  'Set CNMainOFFICE=Server.CreateObject("ADODB.Connection")
  'CNMainOFFICE.open ConnectOffice,"",""
  RSMain.Open "select * from MailXX ",Connect,1,3 '// 新增、修改、或删除
  RSMain.Addnew
  RSMain("MailSJ")=mytime()
  RSMain("MailJB")="普通"
  RSMain("MailFX")=YHMC
  RSMain("MailSX")=MailSX
  RSMain("MailBT")=MailBT
  RSMain("MailNR")=MailNR
  RsMain("MailBM")=UserBM
  RsMain("MailURL")=MailURL
  RsMain("MailBZ")="WEBOA通知"
  RSMain.Update
  MailDM=RSMain("MailDM")
  RSMain.Close
  '---------------------------------------------------------------
  '// 写发信人记录
  if(YHMC<>"")then
     'CNMainOffice.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFLL) values('发信'," & MailDM & ",'"+YHMC+"','发件箱','发件箱',1)")
     CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFLL,SFBM,SFBZ) values('发信'," & MailDM & ",'"+YHMC+"','发通知箱','发通知箱',1,'"+UserBM+"','WEBOA通知')")
  end if
  '---------------------------------------------------------------
  '// 写邮件的收信人记录
  temp="'"+replace(MailSX,",","','")+"'"
  'RSMain.open "select cstr(YHMC) from YHXX where YHMC in ("+temp+") union select cstr(QZYM) from QZXX where QZMC in("+temp+")",Connect,1,1 '// 读取
  RSMain.open "select YHMC from YHXX where YHMC in ("+temp+") union select QZYM from QZXX where QZMC in("+temp+")",Connect,1,1 '// 读取
  temp="*"
  while (not RSMain.Eof)
     temp=temp+","+RSMain(0)
     RSMain.Movenext
     wend
  RSMain.Close
  temp="'"+replace(temp,",","','")+"'"
  CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFBZ) select '收信'," & MailDM & ",YHMC,'收通知箱','收通知箱','WEBOA通知' from YHXX where (YHMC in ("+temp+"))and(YHMC<>'"+YHMC+"')")
  'CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFBZ) select '收信'," & MailDM & ",YHMC,'收通知箱','收通知箱' from YHXX where (YHMC in ("+temp+"))and(YHMC<>'"+YHMC+"')")
  
  '---------------------------------------------------------------
end function

function Mailto_old(RSMain,YHMC,MailSX,MailBT,MailNR)
  '// 发送简单信件
  '---------------------------------------------------------------
  '// 写信件
  RSMain.Open "select * from MailXX ",Connect,1,3 '// 新增、修改、或删除
  RSMain.Addnew
  RSMain("MailSJ")=mytime()
  RSMain("MailJB")="普通"
  RSMain("MailFX")=YHMC
  RSMain("MailSX")=MailSX
  RSMain("MailBT")=MailBT
  RSMain("MailNR")=MailNR
  RSMain.Update
  MailDM=RSMain("MailDM")
  RSMain.Close
  '---------------------------------------------------------------
  '// 写发信人记录
  if(YHMC<>"")then
     CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2,SFLL) values('发信'," & MailDM & ",'"+YHMC+"','发件箱','发件箱',1)")
     end if
  '---------------------------------------------------------------
  '// 写邮件的收信人记录
  temp="'"+replace(MailSX,",","','")+"'"
  'RSMain.open "select cstr(YHMC) from YHXX where YHMC in ("+temp+") union select cstr(QZYM) from QZXX where QZMC in("+temp+")",Connect,1,1 '// 读取
  RSMain.open "select YHMC from YHXX where YHMC in ("+temp+") union select QZYM from QZXX where QZMC in("+temp+")",Connect,1,1 '// 读取
  temp="*"
  while (not RSMain.Eof)
     temp=temp+","+RSMain(0)
     RSMain.Movenext
     wend
  RSMain.Close
  temp="'"+replace(temp,",","','")+"'"
  CNMain.Execute("insert into MailSF (SFZL,SFYJ,SFYM,SFML,SFM2) select '收信'," & MailDM & ",YHMC,'收件箱','收件箱' from YHXX where (YHMC in ("+temp+"))and(YHMC<>'"+YHMC+"')")
  '---------------------------------------------------------------
end function

  '---------------------------------------------------------------
  CoopXX=session("sys.CoopXX")
  CoopDM=fieldget(CoopXX,1,chr(5))
  CoopMC=fieldget(CoopXX,2,chr(5))
  Connect=fieldget(CoopXX,3,chr(5))
  '---------------------------------------------------------------
  UserXX=session("sys.UserXX")
  UserDM=fieldget(UserXX,1,chr(5))  '// 代码
  UserDL=fieldget(UserXX,2,chr(5))  '// 登录名
  UserMC=fieldget(UserXX,3,chr(5))  '// 用户名
  UserBM=fieldget(UserXX,4,chr(5))  '// 部门
  UserZL=fieldget(UserXX,5,chr(5))  '// 种类
  UserP0=session("sys.UserP0")  '// 我的邮箱
  UserP1=session("sys.UserP1") '// 群组
  UserP2=session("sys.UserP2") '// 群组
  UserPZ=session("sys.UserPZ")  '// 配置选项
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
  
  if(UserPZ4<>"标准")then UserPZ4="默认"
  thisfile=Request.ServerVariables("SCRIPT_NAME")
  '---------------------------------------------------------------
  '// StrinfDetect="<SCRIPT language=JavaScript><!--"+chr(13)+chr(10)+"if(parent.logo.window.oa.alt!='企业应用平台') window.location.href='login.asp?page=02';"+chr(13)+chr(10)+"//--></SCRIPT>"+chr(13)+chr(10)
  StringAlarm="严禁在未登录或超时状况下使用信息系统。<br>要重新登录请点<a target=_parent href='login.asp'>这里</a>。"
  StringWind="<center><table border=0 width=100% cellspacing=0 cellpadding=1 bgcolor=#FFFFFF><tr><td><nolayer><iframe name=_edit height=15 width=100% marginwidth=0 marginheight=0 scrolling=no frameborder=0></iframe></nolayer></td></tr></table></center>"
  StringHead="<HTML><HEAD><TITLE>"+CoopMC+"Iasi Information System V2.0</TITLE><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><LINK rel=""stylesheet"" href=""style.css""></HEAD>"
  StringBody="<body bgcolor=#ffffff leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>"
  StringBody2="<body bgcolor=#ffffff leftmargin=5 rightmargin=5 topmargin=5 bottommargin=5>"
  if(UserPZ4="标准")then
     StringWind="<center><table border=1 width=100% cellspacing=0 cellpadding=1 bgcolor=#A0C0E8><tr><td><nolayer><iframe name=_edit height=15 width=100% marginwidth=0 marginheight=0 scrolling=no frameborder=0></iframe></nolayer></td></tr></table></center>"
     StringBody="<body _bgcolor=#FFF9E6 bgcolor=#4488CC leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>"
     StringBody2="<body _bgcolor=#FFF9E6 bgcolor=#4488CC leftmargin=5 rightmargin=5 topmargin=5 bottommargin=5>"
     end if
  '---------------------------------------------------------------
%>
