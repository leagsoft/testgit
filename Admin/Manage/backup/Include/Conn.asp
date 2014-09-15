<%
'Êý¾Ý¿âÁ´½Ó
Dim Conn
Set Conn=Server.CreateObject("Adodb.Connection")
Conn.ConnectionString="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=WEBDATA;Data Source=BENNY"
Conn.Open
%>