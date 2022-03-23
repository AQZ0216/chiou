
<table border="0" cellspacing="0" cellpadding="0">
<col width=120><col width=120><col width=120><col width=120><col width=120>
<tr><td align="center" valign="middle">
	<!-- 人員選項 -->
	<font size=3>人員選項：<br>
		<SELECT name="men_w" onchange="menadd">
		<option selected>請選擇人員</option>
	<%
		for i=1 to writer_no
			response.write "<option value='"&writer_a(i-1)&"'>"&writer_a(i-1)
		next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- 日期選項 -->
	<font size=3>日期選項：<br>
		<SELECT name="date_w" onchange="dateadd">
		<option selected>請選擇日期</option>
	<%
	for i=1 to 60
		Response.Write("<OPTION value=" & wk_date(i-1) & ">" & wk_date(i-1)&"</OPTION>")
	next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- 時間選項 -->
	<font size=3>時間選項：<br>
		<SELECT name="time_w" onchange="timeadd">
		<option selected>請選擇時間</option>
	<%
	for i=1 to 19
		Response.Write("<OPTION value=" & wk_time(i-1) & ">" & wk_time(i-1)&"</OPTION>")
	next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- 地點選項 -->
	<font size=3>地點選項：<br>
		<SELECT name="place_w" onchange="placeadd">
		<option selected>請選擇地點</option>
	<%
		for i=1 to place_no
			response.write "<option value='"&place_a(i-1)&"'>"&place_a(i-1)
		next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- 事件選項 -->
	<font size=3>事件選項：<br>
		<SELECT name="thing_w" onchange="thingadd">
		<option selected>請選擇事件</option>
	<%
		for i=1 to thing_no
			response.write "<option value='"&thing_a(i-1)&"'>"&thing_a(i-1)
		next
	%>
		</font></SELECT>
	</td>

</tr>
</table>
