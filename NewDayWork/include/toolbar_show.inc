
<script language=vbscript>
sub edit_click()
	wk_id=document.form1.wk_id1.value
	location.href="./wk_edit.asp?wk_id="&wk_id
end sub
sub delete_click()
	worker=document.form1.worker1.value
	wk_order=document.form1.wk_ordera.value
	if worker=wk_order then
		ok=msgbox("�O�_�T�w�n�R����ơH",1,"�R��ĵ�i")
		if ok=1 then 
			wk_id=document.form1.wk_id1.value
			location.href="./wk_del_ok.asp?wk_id="&wk_id
		end if
	else
		ok=msgbox("�A���O���u�̡A�L�k�R�������u�@�I�I",0,"���~ĵ�i")
	end if
end sub
sub done_click()
	ok=msgbox("�O�_�T�w�n�����u�@�H",1,"�T�{ĵ�i")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_done_ok.asp?wk_id="&wk_id
	end if
end sub
sub readd_click()
	ok=msgbox("�O�_�T�w�n���s���i�u�@�H",1,"�T�{ĵ�i")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_readd.asp?wk_id="&wk_id
	end if
end sub
sub gpchange_click()
	ok=msgbox("�O�_�T�w�n�ର�M�פu�@�H",1,"�T�{ĵ�i")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_gpchg_special.asp?wk_id="&wk_id
	end if
end sub
</script>
<center>

<table border=0 cellspacing=0 cellpadding=0>
<col span=8 style="width:90px;text-align:center;">
<tr width=720>	
	<td>
		<input type=button name="bkpg" value="�^�W�@��" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="parent.location.href='javascript:history.back()'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="edit" value="�s�פu�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="edit_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="delete" value="�R���u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="delete_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
<% if wk_class="Z" then %>
		<input type=button name="done" value="�����u�@" disabled style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
<% else %>
		<input type=button name="done" value="�����u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="done_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
<% end if %>
	</td>
	<td>
		<input type=button name="readd" value="���s���i" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="readd_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="wkprint" value="�C�L���e" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="window.open('./wkprint_si.asp?wk_id=<%=wk_id%>')"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="gpchange" value="�ର�M��" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="gpchange_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="wkattfile" value="�W�Ǫ���" style="cursor:hand;background-color:'#77FFEE';color:blue;width:100%;" onclick="location.href='./1_ulf_form.asp?wk_id=<%=wk_id%>'"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#77FFEE';">
	</td>
</tr>
</table>
</center>
