//��������ҳ�б�
//������ҳ������,��ǰҳ��,��������
function Tkl_PageListBar(pageCount,CurrentPage,parameter)
{
	if(pageCount<=0){pageCount=1}
	var PageListBarId=0
    for(var i=0;i<=100;i++)
	{		
		if(!eval("window.Tkl_PageListBar"+i))
		{
			PageListBarId=i
			break
		}
	}
    if(parameter!="")
    {
        parameter="&"+parameter
    }
	document.write("<div id=\"Tkl_PageListBar"+PageListBarId+"\">")
    if(CurrentPage<=1){
        document.write("��ҳ|")
    }else{
        document.write("<a href=?CurrentPage=1"+parameter+">��ҳ</a>|")
    }

    if(1<CurrentPage){
        document.write("<a href=?CurrentPage="+(CurrentPage-1)+""+parameter+">��һҳ</a>|")
    }else{
        document.write("��һҳ|")
    }
    document.write(CurrentPage+"/"+pageCount+"ҳ|")
    if(0<CurrentPage && CurrentPage<pageCount){
        document.write("<a href=?CurrentPage="+(CurrentPage+1)+""+parameter+">��һҳ</a>|")
    }else{
        document.write("��һҳ|")
    }
    if(CurrentPage>=pageCount){
        document.write("βҳ")
    }else{
        document.write("<a href=?CurrentPage="+pageCount+""+parameter+">βҳ</a>")
    }
    document.write("&nbsp;<INPUT TYPE=\"text\" size=\"3\" onmouseover=\"this.focus();this.select()\" id=\"Tkl_CurrentPage"+PageListBarId+"\" NAME=\"PGNumber"+PageListBarId+"\" value=\""+CurrentPage+"\" style=\"font-size:9pt;background-color:#f7f7f7;border-left: 1px solid rgb(192,192,192); border-right: 1px solid rgb(192,192,192); border-top: 1px solid rgb(192,192,192); border-bottom: 1px solid rgb(192,192,192)\"><INPUT TYPE=\"button\" value=\"GO\" onclick=\"if(1<=Tkl_CurrentPage"+PageListBarId+".value && Tkl_CurrentPage"+PageListBarId+".value<="+pageCount+"){window.location='?CurrentPage='+Tkl_CurrentPage"+PageListBarId+".value+'"+parameter+"'}\" onmouseover=\"this.focus()\" style=\"font-size:9pt;background-color:#f7f7f7;border-left: 1px solid rgb(192,192,192); border-right: 1px solid rgb(192,192,192); border-top: 1px solid rgb(192,192,192); border-bottom: 1px solid rgb(192,192,192)\"></div>")
}