<table border=0 cellspacing=0 cellpadding=0>
<%	
	col_no=9
	col_no_last=worker_no mod col_no
	if col_no_last=0 then 
		row_no=int(worker_no/col_no)
	else
		row_no=int(worker_no/col_no)+1
	end if
'	col_no_last=worker_no mod col_no
for j=1 to row_no	
%>
<tr>
<%
	for k=1 to col_no
	i=(col_no)*(j-1)+k
	if i > worker_no then
%>
	<td align="center">
<!--	<input class=bar style="cursor:hand;background-color:'#d3d3d3';" type="button" name="worker_s<%=i%>" value="new<%=i%>" onclick="parent.location.href='./wkr_login.asp'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">-->
	<input class=bar style="cursor:hand;background-color:'#d3d3d3';" type="button" name="worker_s<%=i%>" value="¡@" title="<%=i%>" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
<%
	else	
%>
	<td align="center">
	<input class=bar style="cursor:hand;background-color:'#d3d3d3';" type="button" name="worker_s<%=i%>" value="<%=eworker_a(i-1)%>" title="<%=worker_a(i-1)%>" onclick="parent.location.href='./work_main.asp?worker=<%=worker_a(i-1)%>&fp=1'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
<%
	end if
	next
%>
</tr>
<%
next
%>	
</table>
