
<table border=1 cellspacing=0 cellpadding=0>
<col width=120><col width=120><col width=120><col width=120><col width=120><col width=120>
<tr>
	<td align="center" colspan=2><font size=4 color="red"><b>單一工作完成表(<%=wk_group%>)</b></font></td>
	<td align="right">工作編號：</td>
	<td><input type='text' name='wk_id' value='<%=wk_id%>' style="width:100%;" readonly></td>
	<td align="right">工作分類：</td>
	<td><input type='text' name='wk_class' value='<%=wk_class%>' style="width:100%;" readonly></td>
</tr>
<tr>
	<td align="right">公告者：</td>
	<td><input type='text' name='wk_order' value='<%=wk_order%>' style="width:100%;" readonly></td>
	<td align="right">執行日期：</td>
	<td><input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;" readonly></td>
	<td align="right">完成日期：</td>
	<td><input type='text' name='done_date1' value='<%=done_date1%>' style="width:100%;" readonly></td>
</tr>
<tr>
	<td align="right">
	執行人員：
	</td>
	<td colspan=5>
	<input type='text' name='wk_exe' value='<%=wk_exe%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	知會人員：
	</td>
	<td colspan=5>
	<input type='text' name='wk_doer' value='<%=wk_doer%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	完成人員：
	</td>
	<td colspan=5>
	<input type='text' name='wk_checker' value='<%=wk_checker%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	未完成人員：
	</td>
	<td colspan=5>
	<input type='text' name='wk_undoer' value='<%=wk_undoer%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	主旨：
	</td>
	<td colspan=5>
	<input type='text' name='wk_item' value='<%=wk_item%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right" valign="top">
	執行內容：
	</td>
	<td colspan=5>
	<TEXTAREA name="wk_content" rows="10" style="width:100%;" readonly><%=wk_content%></TEXTAREA>
	</td>
</tr>

</table>
