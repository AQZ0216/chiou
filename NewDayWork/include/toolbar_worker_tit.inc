
<table width=1000 border=0 cellspacing=0 cellpadding=0>
<%
   if (worker_no mod 2)=0 then
      tn=worker_no/2
   else
      tn=(worker_no-1)/2
   end if
%>
<tr>	<td align="center">
   <table  border=0 cellspacing=0 cellpadding=0>
   <tr>
<%
	for i=1 to worker_no
%>
   <td align="center">
<%
          if worker=worker_a(i-1) then
            pwkr_id=staff_id_a(i-1)
        %>	
	<input style="cursor:hand;background-color:'#99d700';width:60px;" type="button" name="worker_s<%=i%>" value="<%=worker_a(i-1)%>" onclick="parent.location.href='./work_main.asp?worker=<%=worker_a(i-1)%>'" onmouseover="javascript:this.style.background='#99d700';" onmouseout="javascript:this.style.background='#99d700';">
<% else %>
	<input style="cursor:hand;background-color:'#d3d3d3';width:60px;" type="button" name="worker_s<%=i%>" value="<%=worker_a(i-1)%>" onclick="parent.location.href='./work_main.asp?worker=<%=worker_a(i-1)%>'" onmouseover="javascript:this.style.background='#ffd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
<% end if %>
	</td>
<%
         if i=tn then response.write "</td></tr></table><table  border=0 cellspacing=0 cellpadding=0><tr><td align='center'>"
	next
%>	
</td></tr>

</table>
