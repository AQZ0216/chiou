
<script language=vbscript>
sub edit_onclick
	wk_id=document.form1.wk_id1.value
	location.href="./wk_edit.asp?wk_id="&wk_id
end sub
sub delete_onclick
	worker=document.form1.worker1.value
	wk_order=document.form1.wk_order1.value
	if worker=wk_order then
		ok=msgbox("是否確定要刪除資料？",1,"刪除警告")
		if ok=1 then 
			wk_id=document.form1.wk_id1.value
			location.href="./wk_del_ok.asp?wk_id="&wk_id
		end if
	else
		ok=msgbox("你不是派工者，無法刪除此項工作！！",0,"錯誤警告")
	end if
end sub
sub done_onclick
	ok=msgbox("是否確定要完成工作？",1,"確認警告")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_done_ok.asp?wk_id="&wk_id
	end if
end sub
sub readd_onclick
	ok=msgbox("是否確定要重新公告工作？",1,"確認警告")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_readd.asp?wk_id="&wk_id
	end if
end sub
sub gpchange_onclick
	ok=msgbox("是否確定要轉為專案工作？",1,"確認警告")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_gpchg.asp?wk_id="&wk_id
	end if
end sub
</script>
<center>

<table border=0 cellspacing=0 cellpadding=0>
<col span=7 style="width:100px;text-align:center;">
<tr width=720>	
	<td>
		<input type=button name="bkpg" value="回上一頁" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="parent.location.href='javascript:history.back()'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="edit" value="編修工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';" disabled>
	</td>
	<td>
		<input type=button name="delete" value="刪除工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';" disabled>
	</td>
	<td>
		<input type=button name="done" value="完成工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';" disabled>
	</td>
	<td>
		<input type=button name="readd" value="重新公告" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="wkprint" value="列印內容" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="window.open('./wkprint_si.asp?wk_id=<%=wk_id%>')"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="gpchange" value="轉為專案" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
</tr>
</table>
</center>

