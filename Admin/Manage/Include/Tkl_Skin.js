//鼠标经过
document.onmouseover=ButtonOnMouseOver
document.onmouseout=ButtonOnMouseOut
function ButtonOnMouseOver(){
	try{
		if ((event.srcElement.type=="button")||(event.srcElement.type=="submit")||(event.srcElement.type=="reset"))
		{
			switch(event.srcElement.className)
			{
				case "button01-out" :
					event.srcElement.className="button01-over"
					break
				case "button02-out" :
					event.srcElement.className="button02-over"
					break
			}
		}
	}catch(exception){}
}
//鼠标离开
function ButtonOnMouseOut()
{
	try{
		if ((event.srcElement.type=="button")||(event.srcElement.type=="submit")||(event.srcElement.type=="reset"))
		{
			switch(event.srcElement.className)
			{
				case "button01-over" :
					event.srcElement.className="button01-out"
					break
				case "button02-over" :
					event.srcElement.className="button02-out"
					break
			}
		}
	}catch(exception){}
}