<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
%>
<%
'�]�w�ܼ� 
dim dbConn, rs, nDex, nMonth, nYear, dtDate

' Get the current date
dtDate = Now()

' Set the Month and Year
nMonth = Request("nMonth")
nYear = Request("nYear")
if nMonth = "" then nMonth = Month(dtDate)
if nYear = "" then nYear = Year(dtDate)
select case cint(Weekday(date()))
case 1
	cswday="�P����"
case 2
	cswday="�P���@"
case 3
	cswday="�P���G"
case 4
	cswday="�P���T"
case 5
	cswday="�P���|"
case 6
	cswday="�P����"
case 7
	cswday="�P����"

end select

'�]�w�P���C�� 
bgc1="#F0FFF0" 	'�H����lightyellow 
bgc6="#F0FFF0"	'lightskyblue
bgc7="#F0FFF0"	'lightgreen

' Set the date to the first of the current month
dtDate = DateSerial(nYear, nMonth, 1)


if int(nMonth)<10 then
	strnMonth="0"&cstr(nMonth)
else
	strnMonth=cstr(nMonth)
end if
dcodeym=cstr(nYear)&strnMonth

'�]�wsession("strbackURL")
strbackURL="wk_Calendar_all.asp?nMonth="&nMonth&"&nYear="&nYear
session("strbackURL")=strbackURL

%>
<%
'�]�w�W�@�� 
if nMonth = 1 then 
	pre1month=12
	pre1year=nYear-1
else
	pre1month=nMonth-1
	pre1year=nYear
end if
pre2month=nMonth
pre2year=nYear-1
pre3month=nMonth
pre3year=nYear+1
if nMonth = 12 then 
	pre4month=1
	pre4year=nYear+1
else
	pre4month=nMonth+1
	pre4year=nYear
end if
%>

<!-- �}�Ҹ�Ʈw -->
<%
' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"
%>

<HTML>
<HEAD>
<title>�˪O���D</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{
	margin:10px 0 0 0;		/*��t�W�U���k*/
	font-family:'�s�ө���';		/*�r��*/
	font-size:4mm; 			/*�r��j�p*/
	background-color:'#F0FFF0'; /*�I���C��*/
	}
notetext{
	font-family:'�s�ө���';		/*�r��*/
	font-size:3mm; 			/*�r��j�p*/
	}
daytext{
	font-family:'�s�ө���';		/*�r��*/
	font-size:3mm; 			/*�r��j�p*/
	}
td{
	font-weight:normal;
	}
/*�s��LINK������*/
A:link{color:black}		/*���s�����r���C��*/
A:visited{color:black}	/*���g�s�����r���C��*/
A:active{color:black}	/*�s�����r���C��*/

--></style>

</HEAD>
<body>
<script language=vbscript>
<!--
sub mysel1
	s_month=document.form1.nMonths1.value
	s_Year=document.form1.nYears1.value
	location.href="./wk_calendar_alldone.asp?nMonth="&s_month&"&nYear="&s_Year
end sub
sub mysel2
	s_month=document.form1.nMonths2.value
	s_Year=document.form1.nYears2.value
	location.href="./wk_calendar_alldone.asp?nMonth="&s_month&"&nYear="&s_Year
end sub
-->
</script>
<center>
<form method="post" name="form1" action="">
<table border=0 bgcolor="gray" cellpadding=1 style="width:100%;border-width:0px;width:970px;">
<col style="width:56%;background-color:#F0FFF0;">
<col style="width:11%;background-color:#F0FFF0;">
<col style="width:22%;background-color:#F0FFF0;">
<col style="width:11%;background-color:#F0FFF0;">
<tr style="height:25px">
	<td align=center style="font-size:20px;letter-spacing:1px;" >
		<b><%=worker%>�����u�@����</b>
		<a href="wk_calendar_alldone_pr.asp" target="_blank" style="font-size:3.5mm;letter-spacing:1px;color:red;">[�͵��C�L]</a>
	<td colspan=3 align=center style="font-size:15px;cursor:hand;">
	<a href="wk_calendar_alldone.asp" style="font-size:15px;letter-spacing:1px;font-weight:bold;color:black;">
	���ѬO&nbsp;�褸<%=cstr(Year(date()))%>�~<%=cstr(Month(date()))%>��<%=cstr(Day(date()))%>��&nbsp;<%=cstr(cswday)%> 
	</a>
<tr style="height:25px">
	<td align=center style="font-size:15px;cursor:hand;font-weight:bold;">
	<img SRC='img/list.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='���C�����' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_list_alldone.asp?nMonth=<%=nMonth%>&nYear=<%=nYear%>'">
	&nbsp;
	��ܦ~���G 
	<select name="nYears1" onChange="mysel1" ><%
		' Note: I have set the year to be between 1999 and 2000
		for nDex = 1967 to 2067
			Response.Write "<option value=""" & nDex & """"
			if nDex = CInt(nYear) then Response.Write " selected"
			Response.Write ">" & nDex
		next %></select>
	��ܤ���G
	<select name="nMonths1" onChange="mysel1" ><%
		for nDex = 1 to 12
			Response.Write "<option value=""" & nDex & """"
			if MonthName(nDex) = MonthName(nMonth) then 
				Response.Write " selected"
			end if
			Response.Write ">" & MonthName(nDex)
		next %></select>&nbsp;
	<td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
		<img SRC='img/pre_year.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='�W�@�~' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_alldone.asp?nMonth=<%=pre2month%>&nYear=<%=pre2year%>'">
		<img SRC='img/next_year.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='�U�@�~' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_alldone.asp?nMonth=<%=pre3month%>&nYear=<%=pre3year%>'">
	<td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
		<font style="vertical-align:bottom;font-size:18px;letter-spacing:2px;">�褸<%=cstr(nYear)%>�~<%=cstr(nMonth)%>��</font>
	<td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
		<img SRC='img/pre_month.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='�W�@��' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_alldone.asp?nMonth=<%=pre1month%>&nYear=<%=pre1year%>'">
		<img SRC='img/next_month.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='�U�@��' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_alldone.asp?nMonth=<%=pre4month%>&nYear=<%=pre4year%>'">
</table>
	<table border=0 bgcolor="gray" cellpadding=1 style="width:100%;border-width:0px;width:970px;">
	<col style="width:14%;background-color:<%=bgc7%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc6%>;">
	<tr style="height:45px;">
		<td align=center><font color="black"><b>�P����<br>Sunday</b></font></td>
		<td align=center><font color="black"><b>�P���@<br>Monday</b></font></td>
		<td align=center><font color="black"><b>�P���G<br>Tuesday</b></font></td>
		<td align=center><font color="black"><b>�P���T<br>Wednesday</b></font></td>
		<td align=center><font color="black"><b>�P���|<br>Thursday</b></font></td>
		<td align=center><font color="black"><b>�P����<br>Friday</b></font></td>
		<td align=center><font color="black"><b>�P����<br>Saturday</b></font></td></tr>
	<tr bgcolor="#ffffc0" style="height:60px;">
		<% 
		' Add blank cells until the proper day�W�[�Ů����w��m 
		for nDex = 1 to Weekday(dtDate) - 1
			Response.Write "<td bgcolor=""#c0c0c0"">&nbsp;</td>"
		next
		'�}�l��J��� 
		do
			'����ഫ 
			if int(Day(dtDate))<10 then
				strnDay="0"&cstr(Day(dtDate))
			else
				strnDay=cstr(Day(dtDate))
			end if
			dcodeymd=dcodeym&strnDay
			dcodeymd_a=dtDate
			'�]�w�C�� 
			select case cint(Weekday(dtDate))
			case 1
				bgc=bgc7
			case 7
				bgc=bgc6
			case else 
				bgc=bgc1	
			end select

			'����椺��g��r 
			if int(nYear)=int(Year(date()))and int(nMonth)=int(Month(date()))and int(Day(dtDate))=int(day(date())) then
				Response.Write "<td valign=""top"" bgcolor=""#f0ffae"" >"
			else			
				Response.Write "<td valign=""top"" bgcolor="&bgc&" >"
			end if
					Response.Write "<table border=0 cellpadding=0 style=""width:100%;height:100%;border-width:0px;"">"
					Response.Write "<tr style=""width:100%;height:10px;"">"
					Response.Write "<td align=center style=""width:50%;background-color:#dcdcdc;font-size:13px;"">"
					Response.Write Day(dtDate)
					Response.Write "<td align=center style=""width:50%;background-color:#dcdcdc;"">"
					Response.Write "<font style=""font-size:13px;""><a href='wk_add_date.asp?datecode="&dtDate&"'>�s�W"
					Response.Write "</a></font>"
					'Response.Write "<font style=""text-size:3mm;""><b>&nbsp;"
					'Response.Write "</b></font>"
					Response.Write "<tr><td colspan=2 align=left valign=top>"
			'��J�ƥ��� 
				'Response.Write "<font size=""-1""><b>" &dcodeymd&"</b></font><br>"
				'Response.Write "<font size=""-1""><b>" &totalput&"</b></font><br>"
				'�إ߸�Ʈw�s������	
				set rstObj1=Server.CreateObject("ADODB.Recordset")
				'strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and wk_finisher like '%"&worker&"%' order by wk_item asc , wk_id asc"
				strSQL_show="Select * from " & tb_name & " where (doing_date1 = #"&dcodeymd_a&"# ) and (wk_checker like '%"&worker&"%' ) order by wk_id asc"
				rstObj1.open strSQL_show,conDB,3,1
				totalput=rstObj1.recordcount
			if not rstObj1.EOF then
				rstObj1.Movefirst
				for i = 1 to totalput
					wk_id=rstObj1.fields("wk_id")
					Response.Write "<font style=""font-size:13px;"">"&i&"�B<a href='wk_show.asp?wk_id="&wk_id&"'>" &rstObj1.fields("wk_item") 
					Response.Write "</a></font><br>"
					rstObj1.MoveNext
					if rstObj1.EOF=true then exit for
				next
			else
			end if
				'������ƶ�
				rstObj1.Close
				'���]����ܼ� 
				set rstObj1=Nothing
				
				Response.Write "</table>"
				Response.Write "</td>"
						
				if WeekDay(dtDate) = 7 then 
					Response.Write "</tr>" & vbCrLf & "<tr  bgcolor=""#ffffc0"" style=""height:60px;"">"
				end if
				dtDate = DateAdd("d", 1, dtDate)
				
		loop until (Month(dtDate) <> CInt(nMonth))
		'�����g���� 
			' Add blank cells to fill out the rest of the month if needed�W�[�Ů����w��m
			if Weekday(dtDate) <> 1 then 
				for nDex = Weekday(dtDate) to 7
					Response.Write "<td bgcolor=""#C0C0C0"">&nbsp;</td>"
				next
			end if
			%>
		</tr>
	</table>


</form>
<%

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%> 	
</center>
</body>
</html>
