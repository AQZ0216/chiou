<%@ Language=VBScript CODEPAGE=950 %>
<%
function Weekday_td(wkd) 

	select case cint(wkd)
	case 1
		dcswday="��"
	case 2
		dcswday="�@"
	case 3
		dcswday="�G"
	case 4
		dcswday="�T"
	case 5
		dcswday="�|"
	case 6
		dcswday="��"
	case 7
		dcswday="��"
	end select
	Weekday_td=dcswday
end function

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
bgc1="#ffffff" 	'�H����lightyellow 
bgc6="#ffffff"	'lightskyblue
bgc7="#ffffff"	'lightgreen

' Set the date to the first of the current month
dtDate = DateSerial(nYear, nMonth, 1)


if int(nMonth)<10 then
	strnMonth="0"&cstr(nMonth)
else
	strnMonth=cstr(nMonth)
end if
dcodeym=cstr(nYear)&strnMonth

'�]�wsession("strbackURL")
strbackURL="wk_Calendar_all_email.asp?nMonth="&nMonth&"&nYear="&nYear
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
<!--Ū�J�C�L�˪O�� base_print.css  -->
	<link rel="stylesheet" type="text/css" 
		media="screen" href="./css/base_print.css" title="style_print">
	<link rel="stylesheet" type="text/css" 
		media="print" href="./css/base_print.css" title="style_print">
<style type="text/css"><!--
body{
	margin:10px 0 0 0;		/*��t�W�U���k*/
	font-family:'�з���';		/*�r��*/
	font-size:12.5px; 			/*�r��j�p*/
	background-color:'#F0FFF0'; /*�I���C��*/
	}
notetext{
	font-family:'�з���';		/*�r��*/
	font-size:12.5px; 			/*�r��j�p*/
	}
daytext{
	font-family:'�з���';		/*�r��*/
	font-size:12.5px; 			/*�r��j�p*/
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
	location.href="./wk_calendar_all_pr.asp?nMonth="&s_month&"&nYear="&s_Year
end sub
sub mysel2
	s_month=document.form1.nMonths2.value
	s_Year=document.form1.nYears2.value
	location.href="./wk_calendar_all_pr.asp?nMonth="&s_month&"&nYear="&s_Year
end sub
-->
</script>
<center>
<form method="post" name="form1" action="">
<table border=0 bgcolor="gray" cellpadding=0 style="border-width:0px;width:700px;">
<col style="width:57%;background-color:#ffffff;">
<col style="width:43%;background-color:#ffffff;">
<tr style="height:25px">
	<td align=center style="font-size:20px;letter-spacing:1px;" >
		<b><%=worker%>�������u�@����</b>
	<td align=center style="font-size:15px;cursor:hand;">
		���ѬO&nbsp;�褸<%=cstr(Year(date()))%>�~<%=cstr(Month(date()))%>��<%=cstr(Day(date()))%>��&nbsp;<%=cstr(cswday)%> 
</table>

	<table border=0 bgcolor="gray" cellpadding=0 style="border-width:0px;width:700px;">
	<col style="width:14%;background-color:<%=bgc7%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc1%>;">
	<col style="width:14%;background-color:<%=bgc6%>;">
<!--	
	<tr style="height:45px;">
		<td align=center><font color="black"><b>�P����<br>Sunday</b></font></td>
		<td align=center><font color="black"><b>�P���@<br>Monday</b></font></td>
		<td align=center><font color="black"><b>�P���G<br>Tuesday</b></font></td>
		<td align=center><font color="black"><b>�P���T<br>Wednesday</b></font></td>
		<td align=center><font color="black"><b>�P���|<br>Thursday</b></font></td>
		<td align=center><font color="black"><b>�P����<br>Friday</b></font></td>
		<td align=center><font color="black"><b>�P����<br>Saturday</b></font></td></tr>
--> 
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
			
			weekd1=Weekday(dtDate)
			str_wkd=Weekday_td(weekd1)

			'����椺��g��r 
			if int(nYear)=int(Year(date()))and int(nMonth)=int(Month(date()))and int(Day(dtDate))=int(day(date())) then
				Response.Write "<td valign=""top"" bgcolor=""#f0ffae"" >"
			else			
				Response.Write "<td valign=""top"" bgcolor="&bgc&" >"
			end if
					Response.Write "<table border=0 cellpadding=0 style=""width:100%;height:100%;border-width:0px;"">"
					Response.Write "<tr style=""width:100%;height:10px;"">"
					Response.Write "<td align=center style=""background-color:#dcdcdc;font-size:13px;padding:5 5 5 5;border-width:0 0 1 0;"">"
					Response.Write month(dtDate)&" / "&Day(dtDate)&"�@("&str_wkd&")"
					Response.Write "<tr><td align=left valign=top style=""padding:5 5 5 5;border-width:0px;"" >"
			'��J�ƥ��� 
				'�إ߸�Ʈw�s������	
				set rstObj1=Server.CreateObject("ADODB.Recordset")
				strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and wk_undoer like '%"&worker&"%' order by wk_item asc , wk_id asc"
				rstObj1.open strSQL_show,conDB,3,1
				totalput=rstObj1.recordcount
			if not rstObj1.EOF then
				rstObj1.Movefirst
				for i = 1 to totalput
					wk_id=rstObj1.fields("wk_id")
					Response.Write "<font style=""font-size:13px;"">"&i&"�B" &rstObj1.fields("wk_item") 
					Response.Write "</font><br>"
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
<!-- �s�W�Ƶ����-->
<table border=0 bgcolor="gray" cellpadding=0 style="border-width:0px;width:700px;">
<tr style="height:4cm;">
<td >
<textarea style="width:100%;height:100%;overflow:hidden;">
</textarea>
</td>
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
