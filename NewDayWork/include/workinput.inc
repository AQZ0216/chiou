<%
dim wk_time(48)
dim wk_date(60)
%>
<%
'wk_time=array("04:00","05:00","06:00","06:00","07:00","08:00","09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00","18:00","19:00","20:00","21:00","22:00")
for i=0 to 47
	k=i mod 2
	j=i\2
	if j<10 then 
		j1="0"&cstr(j)
	else
		j1=cstr(j)
	end if
	if k=0 then 
		k1="00"
	else
		k1="30"
	end if
	wk_time(i)=j1&":"&k1
next

for i=1 to 60
	wk_date(i-1)=cstr(date()+i)
next
%>
