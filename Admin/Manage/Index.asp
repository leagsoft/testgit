<!--#Include File="Include/Config.asp"-->
<HTML>
<HEAD>
<TITLE>
... <%=Def_SysTitle&"��Server��"&Request.ServerVariables("SERVER_NAME")&":"&Request.ServerVariables("SERVER_PORT")&"��"%>
　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　
</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="Tsys2003,FuJinFuZhou,ChanGong_Studio">
<meta name="Version" content="Tsys V1.1">
<style>
body {
    margin: 0px;
    padding: 0px;
    border: none;
}
</style>
</HEAD>
<script>
//卦指麼完笥
function getWin()
{
    return window.MainFrame
}

//曜竃狼由
function ExitSys()
{
    getWin().main.location='Login.asp?Work=LogOut&CloseWin=1';
}

function onUnloadSys()
{
    if(confirm('低鳩協曜竃狼由��')){return true}else{return false}
}
</script>
<BODY bgcolor="buttonface">
<iframe frameborder="0" scrolling="no" width="100%" height="100%" src="Main.asp" id="MainFrame"></iframe>
</BODY>
</HTML>
