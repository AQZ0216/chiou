
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<%	
	for i=1 to worker_no
%>
	<td align="center">
	<input style="cursor:hand;" type="button" name="worker_s<%=i%>" value="<%=worker_a(i-1)%>" onclick="parent.location.href='./work_main.asp?worker=<%=worker_a(i-1)%>'">
	</td>
<%
	next
%>	
</tr>
</table>
