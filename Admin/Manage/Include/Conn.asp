<%
'Êý¾Ý¿âÁ´½Ó
 
server.ScriptTimeout =60000
Dim Conn,ConnStr
Set Conn=Server.CreateObject("Adodb.Connection")
ConnStr="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=DataCBRCGD;Data Source=10.100.0.2;Pwd=weboaadmin2004;"
 
Conn.Open ConnStr



Dim Connect,ConnectStr
Set Connect=Server.CreateObject("Adodb.Connection")
ConnectStr="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=OA;Data Source=10.100.0.2;Pwd=weboaadmin2004"
Connect.Open ConnectStr
%>