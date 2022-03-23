
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
if wk_group1="專案工作" then
%>
	<td align="center" colspan=6><font size=3 color="red">註：僅能修改記錄執行內容、知會人員及執行日期</font></td>
<%
else
%>
	<td align="center" colspan=6><font size=3 color="red">註：僅能修改記錄執行內容</font></td>
<%
end if
%>
</tr>

<tr>
	<td align="center" colspan=2 rowspan=2><font size=4 color="red"><b>編修單一工作表</b></font></td>
	<td align="right">工作群組：</td>
	<td><input type='text' name='wk_group' value='<%=wk_group1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
	<td align="right">專案名稱：</td>
	<td>
	<input type=hidden name="wk_pjn" value="<%=wk_pjid%>，<%=wk_pjn%>">
	<input type='text' name='wk_pjn_s' value='<%=wk_pjn%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
<!-- 	<td align="center" colspan=2><font size=4 color="red"><b>編修單一工作表(<%=wk_group1%>)</b></font></td>
 -->
 	<td align="right">工作編號：</td>
	<td><input type='text' name='wk_id' value='<%=wk_id%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
	<td align="right">工作分類：</td>
	<td><input type='text' name='wk_class' value='<%=wk_class1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
</tr>

<tr>
	<td align="right">公告者：</td>
	<td><input type='text' name='wk_order' value='<%=wk_order%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
	<td align="right">公告日期：</td>
	<td><input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
<%
if wk_group1="專案工作" then
%>
	<td align="right"><font color="red"><b>執行日期：</b></font></td>
	<td><input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;"></td>
<%
else
%>
	<td align="right">執行日期：</td>
	<td><input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;"></td>
<%
end if
%>
</tr>
<tr style="background-color:#FFFF33;">
	<td align="right">
	<font color="blue">執行人員：</font>
	</td>
	<td colspan=5>
	<input type='text' name='wk_exe' value='<%=wk_exe1%>' style="width:50%;" readonly title="執行人員請採用右方下拉選單輸入！！！" onkeydown="javascript:if(window.event.keyCode==8) return false;">

	</td>
</tr>
<tr style="background-color:#FFBFFF;">
	<td align="right">
		<font style="background-color:#ddd;text-decoration:none;cursor:hand;color:red;" onclick="addatt_none()" title="清除出席人員資料">[清]</font>
	<font style="color:blue;">出席人員：</font>
	</td>
	<td colspan=5>
	<input type='text' name='wk_att' value='<%=wk_att%>' style="width:50%;"  >
		<SELECT name="attmen_w" onchange="attadd()">
		<option value="" selected>請選擇人員</option>
		<option value="clear" >清除人員</option>
			<option value="全體人員" >全體人員</option>
		<option value="業務全體" >業務全體</option>
		<option value="內勤全體" >內勤全體</option>
	<%
		for i=1 to worker_no
			response.write "<option value='" & worker_a(i-1) & "'>" & worker_a(i-1) &"</option>"
		next
	%>

		</SELECT>

		<SELECT name="attmen_dp" onchange="attadddp()">
			<option value="" selected>部門選擇</option>
			<option value="clear" >清除人員</option>
			<option value="<%=stra_dp01%>" >總經理室</option>
			<option value="<%=stra_dp02%>" >管理部</option>
			<option value="<%=stra_dp03%>" >企劃部</option>
			<option value="<%=stra_dp04%>" >業務部</option>
			<option value="<%=stra_dp05%>" >法務部</option>
			<option value="<%=stra_dp06%>" >財務部</option>
			<option value="<%=stra_dp07%>" >資訊部</option>
			<option value="<%=stra_dp08%>" >建設部</option>
<!--			<option value="<%=stra_dp09%>" >社企</option>-->
			<option value="<%=stra_dp10%>" >我家農業</option>
			<option value="<%=stra_dpa1%>" >業一</option>
			<option value="<%=stra_dpa2%>" >業二</option>
			<option value="<%=stra_dpa3%>" >業Three</option>
			<option value="<%=stra_dpa4%>" >YES</option>
			<option value="<%=stra_dpa5%>" >冒泡業八</option>
		</SELECT>	
				(請輸入必須出席人員)	
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red"><b>知會人員：</b></font>
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
	<input type="button" name="all_sele" value="全部人員" onclick="all_sele_click()">
 	<input type="button" name="all_unsele" value="清除人員" onclick="all_unsele_click()">
	</td>
</tr>
-->

<tr>
	<td align="right">
	完成人員：
	</td>
	<td colspan=5>
	<input type='text' name='wk_checker' value='<%=wk_checker%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right">
	未完成人員：
	</td>
	<td colspan=5>
	<input type='text' name='wk_undoer' value='<%=wk_undoer%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right">
	主旨：
	</td>
	<td colspan=5>
	<input type='text' name='wk_item' value='<%=wk_item%>' style="width:100%;" readonly onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right" valign="top" bgcolor="yellow">
	<font color="red"><b>執行內容：</b></font>
	</td>
	<td colspan=5 bgcolor="yellow">
	<TEXTAREA name="wk_content" rows="10" style="width:100%;" ><%=wk_content%>&#013;<%=worker%>(<%=now()%>)：&#013;</TEXTAREA>
	</td>
</tr>

</table>
<script language=vbscript>
sub attadd()'出席人員
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

sub attadddp()    '==20170606新增===部門選項===出席人員=====
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