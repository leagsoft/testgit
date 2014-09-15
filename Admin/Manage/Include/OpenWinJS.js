var myWin
function OpenWin(Url,w,h)
{       
    if (myWin && !myWin.closed)
	{
		myWin.close();		
	}
    myWin = window.open(Url,"ResWin","resizable,scrollbars,width="+w+",height="+h);
}