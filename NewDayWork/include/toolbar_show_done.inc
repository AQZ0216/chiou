
<script language=vbscript>
sub edit_onclick
	wk_id=document.form1.wk_id1.value
	location.href="./wk_edit.asp?wk_id="&wk_id
end sub
sub delete_onclick
	worker=document.form1.worker1.value
	wk_order=document.form1.wk_order1.value
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
sub done_onclick
	ok=msgbox("�O�_�T�w�n�����u�@�H",1,"�T�{ĵ�i")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_done_ok.asp?wk_id="&wk_id
	end if
end sub
sub readd_onclick
	ok=msgbox("�O�_�T�w�n���s���i�u�@�H",1,"�T�{ĵ�i")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_readd.asp?wk_id="&wk_id
	end if
end sub
sub gpchange_onclick
	ok=msgbox("�O�_�T�w�n�ର�M�פu�@�H",1,"�T�{ĵ�i")
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
		<input type=button name="bkpg" value="�^�W�@��" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="parent.location.href='javascript:history.back()'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="edit" value="�s�פu�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';" disabled>
	</td>
	<td>
		<input type=button name="delete" value="�R���u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';" disabled>
	</td>
	<td>
		<input type=button name="done" value="�����u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';" disabled>
	</td>
	<td>
		<input type=button name="readd" value="���s���i" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="wkprint" value="�C�L���e" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="window.open('./wkprint_si.asp?wk_id=<%=wk_id%>')"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="gpchange" value="�ର�M��" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
</tr>
</table>
</center>
