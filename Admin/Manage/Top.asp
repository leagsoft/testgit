<!--#Include File=Include/Config.asp-->
<html>
<head>
<title>Top.asp</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<link href="Include/ManageStyle.css" rel="stylesheet" type="text/css">
<style></style>
</head>
<SCRIPT LANGUAGE="JavaScript">
<!--
var strUrl="http://"
function GoToUrl()
{
    strUrl=prompt('请输入Url\n注意输入http://',strUrl)
    if(strUrl!=''&&strUrl!=null)
    {
        top.getWin().main.location=strUrl
    }else{
        strUrl="http://"
    }
}
function Tsys_Href()
{
     top.getWin().main.location.reload()
}
//-->
</SCRIPT>
<body bgcolor="#003366" leftmargin="0" topmargin="0" onselectstart="return false" ondragstart="return false" scroll="none">
<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="0">
  <tr> 
    <td width="89%" height="37" align="left" valign="bottom">&nbsp;
      <!--<label id="hidemenuLabel" onclick="top.getWin().HiddenMenu()" title="显示/隐藏" class="MenuItem">隐藏菜单</label> 
      <span class="MenuItem">|</span> <label onclick="window.history.back()" class="MenuItem">后 
      退 <span class="MenuItem">|</span> </label> <label onclick="Tsys_Href()" class="MenuItem">刷 
      新</label> <span class="MenuItem">|</span> <label onclick="window.history.forward()" class="MenuItem">前 
      进</label> <span class="MenuItem">| 
      <label onclick="GoToUrl()" class="MenuItem">浏 览</label>
      | 
      <label class="MenuItem" <%If ConfirmWhenExitNewsSystem Then Response.Write("onclick=""if(confirm('你确定退出本系统？')){top.ExitSys()}""") Else Response.Write("onclick=""top.ExitSys()""") End If%>>关闭系统</label>-->
      </span> 
    <td width="11%" rowspan="2" align="center" valign="bottom"><img src="Images/Manage/CBRCGD.gif"></td>
  </tr>
  <tr>
    <td></tr>
</table>
</body>
</html>