<%
dim iscomment
iscomment=trim(request.QueryString("iscomment"))
%>
<body SCROLLING=no style='MARGIN:0px' <% if iscomment="true" then response.Write "onLoad='getFocus()'" %> >
<link rel="STYLESHEET" type="text/css" href="wbTextBox/edit.css">
<script src="wbTextBox/edit.js" type="text/javascript"></script>
<input type="hidden" name="Body" id="Body" value="" >
<table ID='WBTB_Container' class="WBTB_Body" width=100% height=100% cellpadding=3 cellspacing=0 border=0 >
  <tr id="WBTB_Toolbars"> 
    <td > 
      <table cellpadding=0 cellspacing=0>
        <tr class="yToolbar"> 
          <td> <select language="javascript" class="WBTB_TBGen" id="FontSize" onchange="WBTB_format('fontsize',this[this.selectedIndex].value);">
              <option class="heading" selected>字号 
              <option value="1">1 
              <option value="2">2 
              <option value="3">3 
              <option value="4">4 
              <option value="5">5 
              <option value="6">6 
              <option value="7">7</option>
            </select> </td>
          <td class="WBTB_Btn" TITLE="加粗" LANGUAGE="javascript" onclick="WBTB_format('bold');"> 
            <img class="WBTB_Ico" src="wbTextBox/images/bold.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="斜体" LANGUAGE="javascript" onclick="WBTB_format('italic');"> 
            <img class="WBTB_Ico" src="wbTextBox/images/italic.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="下划线" LANGUAGE="javascript" onclick="WBTB_format('underline');"> 
            <img class="WBTB_Ico" src="wbTextBox/images/underline.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="取消格式" LANGUAGE="javascript" onclick="WBTB_format1('RemoveFormat');"> 
            <img class="WBTB_Ico" src="wbTextBox/images/removeformat.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="左对齐" NAME="Justify" LANGUAGE="javascript" onclick="WBTB_format('justifyleft');"> 
            <img class="WBTB_Ico" src="wbTextBox/images/aleft.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="居中" NAME="Justify" LANGUAGE="javascript" onclick="WBTB_format('justifycenter');"> 
            <img class="WBTB_Ico" src="wbTextBox/images/center.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="右对齐" NAME="Justify" LANGUAGE="javascript" onclick="WBTB_format('justifyright');"> 
            <img class="WBTB_Ico" src="wbTextBox/images/aright.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="字体颜色" LANGUAGE="javascript" onclick="WBTB_foreColor();"> 
            <img class="WBTB_Ico" src="wbTextBox/images/fgcolor.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="字体背景颜色" LANGUAGE="javascript" onclick="WBTB_backColor();"> 
            <img class="WBTB_Ico" src="wbTextBox/images/fbcolor.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="插入超级链接" LANGUAGE="javascript" onclick="WBTB_UserDialog('CreateLink')"> 
            <img class="WBTB_Ico" src="wbTextBox/images/wlink.gif" WIDTH="18" HEIGHT="18" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="去掉超级链接" LANGUAGE="javascript" onclick="WBTB_format1('Unlink');"> 
            <img class="WBTB_Ico" src="wbTextBox/images/unlink.gif" WIDTH="16" HEIGHT="16" unselectable="on"> 
          </td>
          <td class="WBTB_Btn" TITLE="插入表情" LANGUAGE="javascript" onclick="WBTB_foremot()"> 
            <img class="WBTB_Ico" src="wbTextBox/images/smiley.gif" WIDTH="16" HEIGHT="16"> 
          </td>
          <td class="WBTB_Btn" TITLE="插入引用" LANGUAGE="javascript" onclick="WBTB_specialtype('<div class=quote>','</div>')"> 
            <img class="WBTB_Ico" src="wbTextBox/images/quote.gif" WIDTH="16" HEIGHT="16"></td>
			 <td class="WBTB_Btn" TITLE="清除htm代码" LANGUAGE="javascript" onclick="WBTB_CleanAllHtm();"> 
            <img class="WBTB_Ico" src="wbTextBox/images/cleancode.gif" WIDTH="16" HEIGHT="16"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="100%"><input type="hidden" id="richtext" name="richtext"> <iframe class="WBTB_Composition" ID="WBTB_Composition" onblur="WBTB_CopyData('Body');"  MARGINHEIGHT="5" MARGINWIDTH="5" width="100%" height="100%"></iframe> 
    </td>
  </tr>
</table>
<script language="javascript">WBTB_InitDocument('Body','GB2312');</script>
</BODY>

