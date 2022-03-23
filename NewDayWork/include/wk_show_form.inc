<%
function showspace(ztxt)
   if ztxt="" or isnull(ztxt) then
      pztxt="&nbsp;"
   else
      pztxt=ztxt
   end if
   showspace=pztxt
end function
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<tr>
	<td align="center" colspan=2 rowspan=2>
	<font size=4 color="red"><b>顯示單一工作表</b></font>
<% 
if isnull(wk_headline) then wk_headline=0
if cint(wk_headline) > 5 then %>
<img src="./img/gnome_chess.png" width=32 onclick="parent.content.location.href='0_wk_headline_off_20140728.asp?wk_id=<%=wk_id%>'" title="已在跑馬燈訊息中，移出跑馬燈訊息">
<% else %>
<img src="./img/gnome_chess_d.png" width=32 onclick="parent.content.location.href='0_wk_headline_on_20140728.asp?wk_id=<%=wk_id%>'" title="不在跑馬燈訊息中，轉入跑馬燈訊息">
<% end if %>
	</td>
	<td align="right">工作群組：</td>
	<td><!-- <%=showspace(wk_group)%> -->
	<input type='text' name='wk_group' value='<%=wk_group%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
	<td align="right">專案名稱：</td>
	<td><!-- <%=showspace(wk_pjn)%> -->
 	<input type='text' name='wk_pjn' value='<%=wk_pjn%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>

<tr>
	<td align="right">工作編號：</td>
	<td>
	<input type='text' name='wk_id' value='<%=wk_id%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
	<td align="right">工作分類：</td>
	<td>
	<input type='text' name='wk_class' value='<%=wk_class%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>

<tr>
	<td align="right">公告者：</td>
	<td>
	<input type='text' name='wk_order1' value='<%=wk_order%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
	<td align="right">公告日期：</td>
	<td>
	<input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
	<td align="right">執行日期：</td>
	<td>
 	<input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>

<tr>
	<td align="right">
	執行人員：
	</td>
	<td colspan=5>
	<input type='text' name='wk_exe' value='<%=wk_exe%>' style="width:100%;" readonly  onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right" style="background-color:#FFBFFF;">
	出席人員：
	</td>
	<td colspan=5>
 	<input type='text' name='wk_att' value='<%=wk_att%>' style="width:100%;" readonly  onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right">
	知會人員：
	</td>
	<td colspan=5>
 	<input type='text' name='wk_doer' value='<%=wk_doer%>' style="width:100%;" readonly  onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<!--<tr>
	<td align="right">
	完成人員：
	</td>
	<td colspan=5>
 	<input type='text' name='wk_checker' value='<%=wk_checker%>' style="width:100%;" readonly  onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>-->

<!-- <tr>
	<td align="right">
	未完成人員：
	</td>
	<td colspan=5>
 	<input type='text' name='wk_undoer' value='<%=wk_undoer%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr> -->
<tr>
	<td align="right">
	主旨：
	</td>
	<td colspan=5>
 	<input type='text' name='wk_item' value='<%=wk_item%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right" valign="top">
	執行內容：
	</td>
	<td colspan=5>
<%
'	if wk_content="" or isnull(wk_content) then
'	  wk_content_a=wk_content
'	else
'	  wk_content_a=replace(wk_content,chr(13),"<br>")
'	end if
'	response.write  wk_content_a
	%>
	<TEXTAREA name="wk_content" rows="10" style="width:100%;" readonly  onkeydown="javascript:if(window.event.keyCode==8) return false;"><%=wk_content%></TEXTAREA>

 	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">加密文字：</font>
	</td>
	<td colspan=5>
	<input type='text' name='str_pwd' value='<%=pwk_password%>' style="width:100px;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
</table>
