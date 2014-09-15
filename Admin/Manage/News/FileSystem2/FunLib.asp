<%
'#######¹ýÂËÄ¿Â¼
Function FilterPath(strPath)
    strPath=Replace(Trim(strPath),"../","")
    strPath=Replace(Trim(strPath),"..\","")
    strPath=Replace(strPath,"\..","")
    strPath=Replace(strPath,"/..","")
    FilterPath=strPath
End Function
%>