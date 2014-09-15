<!--#include file="Include/Conn.asp" -->
<!--#include file="Include/Tkl_SYSProedomClass.asp" -->
<!--#include file="Include/Config.asp" -->
<!-- #include file="Include/ClassList_Fun.asp" -->
<!--#include file="Include/CreateFile_Fun.asp" -->
<!--#include file="Include/Tkl_StringClass.asp" -->
<%
Dim SysAdmin
Set SysAdmin=New SYSProedom_Class
If Not SysAdmin.Logined Then
    Response.Redirect("Login.asp")
End If
%>
<html>
<head>
<title>News_CreateHtml.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link href="Include/ManageStyle.css" rel="stylesheet" type="text/css">
<script src="Include/Tkl_Skin.js"></script>
</head>
<%
Dim I
Select Case Request("Work")
    Case "CreateFile"
        Call CreateFile()
End Select
%>
<body bgcolor="#FFFFFF">
<form name="form1" method="post" action="?Work=CreateFile" onSubmit="return checkForm(this)">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ContentTabBg">
    <tr> 
      <td height="16" colspan="2" align="center" class="BarTitleBg">���ɾ�̬ҳ��</td>
    </tr>
    <tr Id="tr1"> 
      <td height="16" colspan="2" bgcolor="#FFFFCC" id="td1"><label for="SelType1"> 
        <input name="SelType1" type="checkbox" id="SelType1" onClick="changeSel(1)" value="1">
        <strong>����һ��ʱ���ڵ�������Դ</strong></label> </td>
    </tr>
    <tr id="tr1_1"> 
      <td width="17%" height="7" align="right" bgcolor="#f6f6f6">����:</td>
      <td width="83%" bgcolor="#FFFFFF"> 
        <select name="TimeType" class="Input" id="TimeType">
          <option value="upTime" selected>������ʱ��</option>
          <option value="AddTime">��Դ����ʱ��</option>
        </select>
      </td>
    </tr>
    <tr id="tr1_2"> 
      <td height="8" align="right" bgcolor="#f6f6f6">��ʼʱ��:</td>
      <td width="83%" bgcolor="#FFFFFF"> 
        <select name="startYear" class="Input">
          <%For I=Year(Now) To 1900 Step -1%>
          <option value="<%=I%>"><%=I%></option>
          <%Next%>
        </select>
        - 
        <select name="startMonth" class="Input">
          <%For I=1 To 12%>
          <option value="<%=I%>" <%If I=Month(Now) Then Response.Write("selected") End If%>><%=I%></option>
          <%Next%>
        </select>
        - 
        <select name="startDay" class="Input">
          <%For I=1 To 31%>
          <option value="<%=I%>" <%If I=Day(Now) Then Response.Write("selected") End If%>><%=I%></option>
          <%Next%>
        </select>
        ( 
        <select name="startHour" class="Input" id="startHour">
          <option value="1" selected>1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
          <option value="6">6</option>
          <option value="7">7</option>
          <option value="8">8</option>
          <option value="9">9</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23">23</option>
          <option value="24">24</option>
        </select>
        ʱ 
        <select name="startMin" class="Input" id="startMin">
          <option value="1" selected>1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
          <option value="6">6</option>
          <option value="7">7</option>
          <option value="8">8</option>
          <option value="9">9</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23">23</option>
          <option value="24">24</option>
          <option value="25">25</option>
          <option value="26">26</option>
          <option value="27">27</option>
          <option value="28">28</option>
          <option value="29">29</option>
          <option value="30">30</option>
          <option value="31">31</option>
          <option value="32">32</option>
          <option value="33">33</option>
          <option value="34">34</option>
          <option value="35">35</option>
          <option value="36">36</option>
          <option value="37">37</option>
          <option value="38">38</option>
          <option value="39">39</option>
          <option value="40">40</option>
          <option value="41">41</option>
          <option value="42">42</option>
          <option value="43">43</option>
          <option value="44">44</option>
          <option value="45">45</option>
          <option value="46">46</option>
          <option value="47">47</option>
          <option value="48">48</option>
          <option value="49">49</option>
          <option value="50">50</option>
          <option value="51">51</option>
          <option value="52">52</option>
          <option value="53">53</option>
          <option value="54">54</option>
          <option value="55">55</option>
          <option value="56">56</option>
          <option value="57">57</option>
          <option value="58">58</option>
          <option value="59">59</option>
          <option value="60">60</option>
        </select>
        ��)</td>
    </tr>
    <tr id="tr1_3"> 
      <td height="16" align="right" bgcolor="#f6f6f6">����ʱ��:</td>
      <td bgcolor="#FFFFFF"> 
        <select name="EndYear" class="Input">
          <%For I=Year(Now) To 1900 Step -1%>
          <option value="<%=I%>"><%=I%></option>
          <%Next%>
        </select>
        - 
        <select name="EndMonth" class="Input">
          <%For I=1 To 12%>
          <option value="<%=I%>" <%If I=Month(Now) Then Response.Write("selected") End If%>><%=I%></option>
          <%Next%>
        </select>
        - 
        <select name="EndDay" class="Input">
          <%For I=1 To 31%>
          <option value="<%=I%>" <%If I=Day(Now) Then Response.Write("selected") End If%>><%=I%></option>
          <%Next%>
        </select>
        ( 
        <select name="endHour" class="Input" id="endHour">
          <option value="1">1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
          <option value="6">6</option>
          <option value="7">7</option>
          <option value="8">8</option>
          <option value="9">9</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23" selected>23</option>
        </select>
        ʱ 
        <select name="endMin" class="Input" id="select3">
          <option value="1">1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
          <option value="6">6</option>
          <option value="7">7</option>
          <option value="8">8</option>
          <option value="9">9</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23">23</option>
          <option value="24">24</option>
          <option value="25">25</option>
          <option value="26">26</option>
          <option value="27">27</option>
          <option value="28">28</option>
          <option value="29">29</option>
          <option value="30">30</option>
          <option value="31">31</option>
          <option value="32">32</option>
          <option value="33">33</option>
          <option value="34">34</option>
          <option value="35">35</option>
          <option value="36">36</option>
          <option value="37">37</option>
          <option value="38">38</option>
          <option value="39">39</option>
          <option value="40">40</option>
          <option value="41">41</option>
          <option value="42">42</option>
          <option value="43">43</option>
          <option value="44">44</option>
          <option value="45">45</option>
          <option value="46">46</option>
          <option value="47">47</option>
          <option value="48">48</option>
          <option value="49">49</option>
          <option value="50">50</option>
          <option value="51">51</option>
          <option value="52">52</option>
          <option value="53">53</option>
          <option value="54">54</option>
          <option value="55">55</option>
          <option value="56">56</option>
          <option value="57">57</option>
          <option value="58">58</option>
          <option value="59" selected>59</option>
        </select>
        ��)</td>
    </tr>
    <tr Id="tr2"> 
      <td height="16" colspan="2" bgcolor="#FFFFFF" id="td2"> <label for="SelType2"> 
        <input name="SelType2" type="checkbox" id="SelType2" value="1" onClick="changeSel(2)">
        <strong>����ָ����Ŀ����Դ</strong></label></td>
    </tr>
    <tr  id="tr2_1"> 
      <td height="7" align="right" bgcolor="#f6f6f6" valign="top">��ѡ����Ŀ:</td>
      <td bgcolor="#FFFFFF"> 
        <script language="JavaScript" src="Include/Tkl_ClassTree.js" type="text/JavaScript"></script>
        <script>
      var root
      root=CreateRoot("myTree","��ѡ��һ�����")
      <%Call ViewTree(0)%>
      </script>
      </td>
    </tr>
    <tr id="tr1_4"> 
      <td height="3" bgcolor="#FFFFFF"><strong> 
        <input name="SelType22" type="checkbox" id="SelType22" value="1" onClick="changeSel(2)" disabled>
        ����ѡ��:</strong></td>
      <td height="3" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <tr id="tr1_4"> 
      <td height="1" bgcolor="#FFFFFF" align="right">��Դ��Χ:</td>
      <td height="1" bgcolor="#FFFFFF"> <label for="ResType1"> 
        <input type="radio" name="ResType" id="ResType1" value="1" checked>
        δ�����������ɵ���Դ</label><label for="ResType2"> 
        <input type="radio" name="ResType" id="ResType2" value="2">
        ��δ����������Դ</label> </td>
    </tr>
    <tr id="tr1_4"> 
      <td height="0" bgcolor="#FFFFFF" align="right">�޶���Ŀ:</td>
      <td height="0" bgcolor="#FFFFFF"> 
        <input type="text" name="TopNum" size="8" value="100" class="Input" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">
        ����Դ (Ϊ'0'���ʾ������Ŀ) </td>
    </tr>
    <tr id="tr1_4">
      <td height="0" bgcolor="#FFFFFF" align="right">���ɱ���:</td>
      <td height="0" bgcolor="#FFFFFF">
        <input type="checkbox" name="CreateReport" value="1">
        ��(�����ɵ���Դ�ϴ󲻽��鹴ѡ����)</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td colspan="2"> 
        <input name="SelType" type="hidden" id="SelType">
        <script>
form1.SelType1.click();
function changeSel(id)
{    
    if(id==1){
        form1.SelType1.checked=true;
        form1.SelType2.checked=false;
        form1.SelType.value=1;
        td1.bgColor="#FFFFCC";
        td2.bgColor="#FFFFff";
        tr1_1.style.display="";
        tr1_2.style.display="";
        tr1_3.style.display="";
        tr2_1.style.display="none";
    }
    if(id==2){
        form1.SelType1.checked=false;
        form1.SelType2.checked=true;
        form1.SelType.value=2;
        td2.bgColor="#FFFFCC";
        td1.bgColor="#FFFFff";
        tr1_1.style.display="none";
        tr1_2.style.display="none";
        tr1_3.style.display="none";
        tr2_1.style.display="";
    }
}
function checkForm(obj){
    if(form1.SelType.value==2){
    }
    return true;    
}
</script>
        <input name="Submit" type="submit" class="button01-out" value="ȷ  ��">
        <input name="Submit2" type="reset" class="button01-out" value="��  ԭ">
        <input name="Submit3" type="button" class="button01-out" value="��  ��" onClick="window.history.back();">
      </td>
    </tr>
  </table>
</form>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2" bgcolor="#f6f6f6" style="cursor:hand" onClick="if(HelpTab.style.display=='none'){HelpTab.style.display='';window.scrollTo(window.pageXOffset,2000);}else{HelpTab.style.display='none'}">&nbsp;<img src="Images/Manage/why.gif" width="14" height="14"> 
      ::Help::</td>
  </tr>
  <tr Id="HelpTab" style="display:none"> 
    <td width="2%">&nbsp;</td>
    <td width="98%" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td>������ָ����Ŀ����Դ</td>
        </tr>
        <tr> 
          <td> 
            <p>����ϵͳֻ���ɵ�ǰ��ѡ��Ŀ��������Դ����������Ŀ��</p>
          </td>
        </tr>
        <tr> 
          <td> 
            <p>��<FONT COLOR="#339900">ϵͳ�涨�����������Դ��������</FONT></p>
          </td>
        </tr>
        <tr> 
          <td>����1.��δ��ˡ�����Դ��2.������վ���ڵ���Դ</td>
        </tr>
        <tr> 
          <td>�����ɽ���</td>
        </tr>
        <tr> 
          <td>����������Դ���ɽ�Ƶ���Է�����Ӳ�̽���I/0��������ռ�ô���ϵͳ��Դ������뾡���ܵ���СҪ���ɵ���Դ��Χ���Ӷ���������ٶȣ����������������</td>
        </tr>
        <tr>
          <td>����</td>
        </tr>
      </table>
      <a name="Help"></a></td>
  </tr>
</table>
</body>
</html>
<%
Sub CreateFile()
	If Not SysAdmin.CreateNewsFile Then
		Response.Write("<script>alert(""<����ʧ��>\n���Ȩ�޲���"& SoftCopyright_Script &""");window.history.back();</script>")
		Response.End()
	End If
	
    Dim Sql
    Dim Rs
    Dim startTime,endTime
    Dim def_Server_ScriptTimeOut
        def_Server_ScriptTimeOut=Server.ScriptTimeOut
    Dim CreateFileStartTime
        CreateFileStartTime=Now()
    Server.ScriptTimeOut=CreateNewsFiles_ScriptTimeOut

    startTime=Request("startMonth") &"/" & Request("startDay") & "/" & Request("startYear") &" "&Request("startHour")&":"&Request("startMin")&":00"
    endTime=Request("endMonth") &"/" & Request("endDay") & "/" & Request("endYear") &" "&Request("endHour")&":"&Request("endMin")&":00"

	Dim TimeType
        TimeType=Request("TimeType")
    If CLng(Request("SelType"))=1 Then
        '����ָ��ʱ�䷶Χ����
        Sql="Select {TopNum} ID,Title,ClassTitle From view_NewsInfo Where #"&startTime&"#<"&TimeType&" And "&TimeType&"<#"&endTime&"# {ResType} Order By Class,Id DESC"
    Else
        '����ָ����������
        If CLng(Request("SelType"))=2 Then
            Dim selItemList
                selItemList=Request("chkBoxItme")
            If selItemList="" Then
                Response.Write("<script>alert(""��ѡ������һ��[��Դ����]"");window.history.back();</script>")
                Response.End
            End If
            Sql="Select {TopNum} ID,Title,ClassTitle From view_NewsInfo Where Class In ("&selItemList&") {ResType} ORDER BY Class,Id DESC"
        End If
    End If
	'��ѡ��
	Select Case Request("ResType")
		Case "1"
			Sql=Replace(Sql,"{ResType}","")
		Case "2"
			Sql=Replace(Sql,"{ResType}","And Created=0")
	End Select
	If IsNumeric(Request("TopNum")) Or Request("TopNum")="0" Then
		Sql=Replace(Sql,"{TopNum}","Top " & Request("TopNum"))		
	Else
		Sql=Replace(Sql,"{TopNum}","")
	End If

    Dim Count
    Count=0
    Set Rs=Conn.ExeCute(Sql)

    '���ģ�建��������
    Session("buffer_NewsTemplate_ClassId")=""
    Session("buffer_NewsTemplate")=""
	Dim CreateReport
		CreateReport=CBool(Request("CreateReport"))
	If CreateReport Then
		Dim Fso1,Fle1
		Set Fso1=Server.CreateObject("Scripting.FileSystemObject")
		Set Fle1=Fso1.OpenTextFile(Server.MapPath("./CreateReport.htm"),2,True)
		Fle1.Writeline("<html><head><title>Tsys��Դ���ɽ������["&Date()&"]</title></head><style>boyd{font-size:9pt}</style><body>"&vbCrLf)
	End If
    While Not Rs.Eof
		'������Դ
        If UsedTemplate_CreateFile(Rs("Id")) Then
			If CreateReport Then
				Fle1.Writeline("<b>"&Count+1&"</b>. ["&Rs("ClassTitle")&"]"&Rs("Title")&" >> �ɹ�<BR>")
			End If
            Count=Count+1
		Else
			If CreateReport Then
				Fle1.Writeline("<b>"&Count+1&"</b>. ["&Rs("ClassTitle")&"]"&Rs("Title")&" >> <b>ʧ��</b><BR>")
			End If
        End If
        Rs.MoveNext
    Wend
	If CreateReport Then
		Fle1.Writeline("<hr height=1>�ܹ����ɣ�<b>"&Count&"<b>������Դ,������ʱ�䣺" & Now())
		Fle1.Writeline("<br>�뼱ʱ���汨���´����ɽ����Ǵ˱���</body></html>")
		Fle1.Close
	End If

    Rs.Close
    Set Rs=Nothing

    '�ָ�ԭIISĬ�ϵĽű���ʱʱ��
    Server.ScriptTimeOut=def_Server_ScriptTimeOut
	If CreateReport Then
	    Response.Write("<script>window.open(""CreateReport.htm"")</script>")		
	End If
    Response.Write("<script>alert(""<�����ɹ�>\n�ļ��������,\nϵͳ��������["&Count&"]����̬��Դ\n������ʱ��(��):"&DateDiff("s",CreateFileStartTime,Now()) & SoftCopyright_Script &""");window.history.back();</script>")
    Response.End
End Sub

Sub ViewTree(ParentId)
    Dim Sql
        Sql="Select * From ClassList Where Parent="&ParentId
    Dim Rs
    Set Rs=Conn.ExeCute(Sql)
    While Not Rs.Eof
        If Rs("Parent")=0 Then
            Response.Write "root.CreateNode("&Rs("Id")&",-1,""<INPUT TYPE=\""checkbox\"" NAME=\""chkBoxItme\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        Else
            Response.Write "root.CreateNode("&Rs("Id")&","&Rs("Parent")&",""<INPUT TYPE=\""checkbox\"" NAME=\""chkBoxItme\"" value=\"""&Rs("Id")&"\"">"&Rs("Title")&""")" & vbCrLf
        End If
        ViewTree(Rs("Id"))
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Sub
%>