
<!--<table border=1 cellspacing=0 cellpadding=0>
<col width=120><col width=120><col width=120><col width=120><col width=120><col width=120>
-->
<table border=1 cellspacing=0 cellpadding=0>
<col width=150>
<col width=180>
<col width=100>
<col width=180>
<col width=100>
<col width=180>
<tr>
<%
if wk_group1="�M�פu�@" then
%>
	<td align="center" colspan=6><font size=3 color="red">���G�ȯ�ק�O�����椺�e�B���|�H���ΰ�����</font></td>
<%
else
%>
	<td align="center" colspan=6><font size=3 color="red">���G�ȯ�ק�O�����椺�e</font></td>
<%
end if
%>
</tr>

<tr>
	<td align="center" colspan=2 rowspan=2><font size=4 color="red"><b>�s�׳�@�u�@��</b></font></td>
	<td align="right">�u�@�s�աG</td>
	<td><input type='text' name='wk_group' value='<%=wk_group1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
	<td align="right">�M�צW�١G</td>
	<td>
	<input type=hidden name="wk_pjn" value="<%=wk_pjid%>�A<%=wk_pjn%>">
	<input type='text' name='wk_pjn_s' value='<%=wk_pjn%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
<!-- 	<td align="center" colspan=2><font size=4 color="red"><b>�s�׳�@�u�@��(<%=wk_group1%>)</b></font></td>
 -->
 	<td align="right">�u�@�s���G</td>
	<td><input type='text' name='wk_id' value='<%=wk_id%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
	<td align="right">�u�@�����G</td>
	<td><input type='text' name='wk_class' value='<%=wk_class1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
</tr>

<tr>
	<td align="right">���i�̡G</td>
	<td><input type='text' name='wk_order' value='<%=wk_order%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
	<td align="right">���i����G</td>
	<td><input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
<%
if wk_group1="�M�פu�@" then
%>
	<td align="right"><font color="red"><b>�������G</b></font></td>
	<td><input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;"></td>
<%
else
%>
	<td align="right">�������G</td>
	<td><input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
<%
end if
%>
</tr>
<tr style="background-color:#FFFF33;">
	<td align="right">
	<font color="blue">����H���G</font>
	</td>
	<td colspan=5>
	<input type='text' name='wk_exe' value='<%=wk_exe1%>' style="width:50%;" readonly title="����H���бĥΥk��U�Կ���J�I�I�I" onkeydown="javascript:if(window.event.keyCode==8) return false;">

	</td>
</tr>
<tr style="background-color:#FFBFFF;">
	<td align="right">
		<font style="background-color:#ddd;text-decoration:none;cursor:hand;color:red;" onclick="addatt_none()" title="�M���X�u�H�����">[�M]</font>
	<font style="color:blue;">�X�u�H���G</font>
	</td>
	<td colspan=5>
	<input type='text' name='wk_att' value='<%=wk_att%>' style="width:50%;"  >
		<SELECT name="attmen_w" onchange="attadd()">
		<option value="" selected>�п�ܤH��</option>
		<option value="clear" >�M���H��</option>
			<option value="����H��" >����H��</option>
		<option value="�~�ȥ���" >�~�ȥ���</option>
		<option value="���ԥ���" >���ԥ���</option>
	<%
		for i=1 to worker_no
			response.write "<option value='" & worker_a(i-1) & "'>" & worker_a(i-1) &"</option>"
		next
	%>

		</SELECT>

		<SELECT name="attmen_dp" onchange="attadddp()">
			<option value="" selected>�������</option>
			<option value="clear" >�M���H��</option>
			<option value="<%=stra_dp01%>" >�`�g�z��</option>
			<option value="<%=stra_dp02%>" >�޲z��</option>
			<option value="<%=stra_dp03%>" >������</option>
			<option value="<%=stra_dp04%>" >�~�ȳ�</option>
			<option value="<%=stra_dp05%>" >�k�ȳ�</option>
			<option value="<%=stra_dp06%>" >�]�ȳ�</option>
			<option value="<%=stra_dp07%>" >��T��</option>
			<option value="<%=stra_dp08%>" >�س]��</option>
<!--			<option value="<%=stra_dp09%>" >����</option>-->
			<option value="<%=stra_dp10%>" >�ڮa�A�~</option>
			<option value="<%=stra_dpa1%>" >�~�@</option>
			<option value="<%=stra_dpa2%>" >�~�G</option>
			<option value="<%=stra_dpa3%>" >�~Three</option>
			<option value="<%=stra_dpa4%>" >YES</option>
			<option value="<%=stra_dpa5%>" >�_�w�~�K</option>
		</SELECT>	
				(�п�J�����X�u�H��)	
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red"><b>���|�H���G</b></font>
	</td>
	<td colspan=5>
	<input type='text' name='wk_doer' value='<%=wk_doer%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>

<!--
<tr>
	<td colspan=6 align="center">
		<table border=0 cellspacing=0 cellpadding=0 style="" >
			<tr style="text-align:center;">
<%	
	for i=1 to worker_no
%>
<td>
	<input type="button"  name="worker_se<%=i%>" value="<%=worker_a(i-1)%>" onclick="worker_se<%=i%>_click()">
</td>
<%
	next
%>	
		</tr>
		</table>
	<input type="button" name="all_sele" value="�����H��" onclick="all_sele_click()">
 	<input type="button" name="all_unsele" value="�M���H��" onclick="all_unsele_click()">
	</td>
</tr>
-->

<tr>
	<td align="right">
	�����H���G
	</td>
	<td colspan=5>
	<input type='text' name='wk_checker' value='<%=wk_checker%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right">
	�������H���G
	</td>
	<td colspan=5>
	<input type='text' name='wk_undoer' value='<%=wk_undoer%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right">
	�D���G
	</td>
	<td colspan=5>
	<input type='text' name='wk_item' value='<%=wk_item%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right" valign="top" bgcolor="yellow">
	<font color="red"><b>���椺�e�G</b></font>
	</td>
	<td colspan=5 bgcolor="yellow">
	<TEXTAREA name="wk_content" rows="10" style="width:100%;" ><%=wk_content%>&#013;<%=worker%>(<%=now()%>)�G&#013;</TEXTAREA>
	</td>
</tr>

</table>
<script language=vbscript>
sub attadd()'�X�u�H��
  if document.form1.attmen_w.value="clear" then
   document.form1.wk_att.value=""
  else
	if document.form1.wk_att.value="" then
		document.form1.wk_att.value=document.form1.attmen_w.value
	else
         if instr(1,document.form1.wk_att.value,document.form1.attmen_w.value,1)>0 then
            document.form1.wk_att.value=replace(document.form1.wk_att.value,document.form1.attmen_w.value,"")
            document.form1.wk_att.value=replace(document.form1.wk_att.value,",,",",")
            if left(document.form1.wk_att.value,1)="," then document.form1.wk_att.value=right(document.form1.wk_att.value,len(document.form1.wk_att.value)-1)
            if right(document.form1.wk_att.value,1)="," then document.form1.wk_att.value=left(document.form1.wk_att.value,len(document.form1.wk_att.value)-1)
         else
		document.form1.wk_att.value=document.form1.wk_att.value & "," & document.form1.attmen_w.value
         end if
	end if
  end if
	document.form1.attmen_w.value=""
end sub	

sub attadddp()    '==20170606�s�W===�����ﶵ===�X�u�H��=====
  if document.form1.attmen_dp.value="clear" then
   document.form1.wk_att.value=""
  else
	if document.form1.wk_att.value="" then
		document.form1.wk_att.value=document.form1.attmen_dp.value
	else
         if instr(1,document.form1.wk_att.value,document.form1.attmen_dp.value,1)>0 then
            document.form1.wk_att.value=replace(document.form1.wk_att.value,document.form1.attmen_dp.value,"")
            document.form1.wk_att.value=replace(document.form1.wk_att.value,",,",",")
            if left(document.form1.wk_att.value,1)="," then document.form1.wk_att.value=right(document.form1.wk_att.value,len(document.form1.wk_att.value)-1)
            if right(document.form1.wk_att.value,1)="," then document.form1.wk_att.value=left(document.form1.wk_att.value,len(document.form1.wk_att.value)-1)
         else
		document.form1.wk_att.value=document.form1.wk_att.value & "," & document.form1.attmen_dp.value
         end if
	end if
  end if
	document.form1.attmen_dp.value=""
end sub

sub addatt_none()
    document.form1.wk_att.value=str_old
end sub
</script>