//---------------------------------------------------------------------
//---------------------------------------------------------------------
//
//							### 数据库专用无限层树 ###
//
//---------------------------------------------------------------------
//---------------------------------------------------------------------
//生成根节点
//	参数：
//		mId:树根的唯一ID
//		Title:显示文字

function CreateRoot(mId,Title)
{
	var TreeId
		TreeId=mId
	if(eval("window.TMNode_"+mId))
	{
		alert("<生成树错误>\n根节点的Id已存在")
		return null
	}
	var str = ""
	str+="<table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\" class=\"TMStyle\">"
	str+="<tr>"
	str+="<td colspan=\"2\"><span style=\"cursor:hand\" onclick=\"if(TMRoot_"+mId+"_Tr.style.display==''){TMRoot_"+mId+"_Tr.style.display='none'}else{TMRoot_"+mId+"_Tr.style.display=''}\">"+Title+"</span></td>"
	str+="</tr>"
	str+="<tr style=\"display:none\" id=\"TMRoot_"+mId+"_Tr\">"
	str+="<td width=\"1%\" align=\"right\"  valign=\"top\">&nbsp;&nbsp;&nbsp;</td>"
	str+="<td width=\"99%\" id=\"TMNode_"+mId+"\"></td>"
	str+="</tr>"
	str+="</table>"
	document.write(str);

	//生成根的子节点
	//	参数：
	//		Id:当前节点的唯一ID
	//		pId:父节点的唯一ID（若为-1,表示其父节点为根Root）
	//		Title:显示文字
	this.CreateNode=function (Id,pId,Title){
		var pNode=null
		if(pId==-1){
			pNode=eval("window.TMNode_"+mId)
		}else{
			pNode=eval("window.TMNode_"+mId+"_"+pId)
		}
		if(pNode!=null)
		{
			var str = ""
			str+="<table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">"
			str+="<tr> "
			str+="<td colspan=\"2\"><span style=\"cursor:hand\" onclick=\"exNode(eval('window.TMNode_"+mId+"_"+Id+"_tr'),eval('window.TMNode_"+mId+"_"+Id+"'))\">"+Title+"</span></td>"
			str+="</tr>"
			str+="<tr style=\"display:none\" id=\"TMNode_"+mId+"_"+Id+"_tr\">"
			str+="<td width=\"1%\" align=\"right\">&nbsp;&nbsp;&nbsp;</td>"
			str+="<td width=\"99%\" id=\"TMNode_"+mId+"_"+Id+"\"></td>"
			str+="</tr>"
			str+="</table>"
			pNode.innerHTML+=str
		}
	}
	return this
}

function exNode(objtr,objNode)
{
	if(objtr&&objNode)
	{		
		if(objNode.innerHTML!="")
		{
			if(objtr.style.display=="")
			{
				objtr.style.display="none";
			}else{
				objtr.style.display="";
			}
		}
	}
}