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
strbackURL="wk_Calendar_list_alldone_pr.asp?nMonth="&nMonth&"&nYear="&nYear
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
	font-size:4mm; 			/*�r��j�p*/
	background-color:'#F0FFF0'; /*�I���C��*/
	}
notetext{
	font-family:'�з���';		/*�r��*/
	font-size:3mm; 			/*�r��j�p*/
	}
daytext{
	font-family:'�з���';		/*�r��*/
	font-size:3mm; 			/*�r��j�p*/
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
	location.href="./wk_calendar_list_alldone_pr.asp?nMonth="&s_month&"&nYear="&s_Year
end sub
sub mysel2
	s_month=document.form1.nMonths2.value
	s_Year=document.form1.nYears2.value
	location.href="./wk_calendar_list_alldone_pr.asp?nMonth="&s_month&"&nYear="&s_Year
end sub
-->
</script>
<center>
<form method="post" name="form1" action="">
<table border=0 bgcolor="gray" cellpadding=1 style="border-width:0px;width:700px;">
<col style="width:56%;background-color:#F0FFF0;">
<col style="width:11%;background-color:#F0FFF0;">
<col style="width:22%;background-color:#F0FFF0;">
<col style="width:11%;background-color:#F0FFF0;">
<tr style="height:25px">
	<td align=center style="font-size:20px;letter-spacing:1px;" >
		<b><%=worker%>�����u�@����</b>
	<td colspan=3 align=center style="font-size:15px;cursor:hand;">
	<a href="wk_calendar_list_alldone_pr.asp" style="font-size:15px;letter-spacing:1px;font-weight:bold;color:black;">
	���ѬO&nbsp;�褸<%=cstr(Year(date()))%>�~<%=cstr(Month(date()))%>��<%=cstr(Day(date()))%>��&nbsp;<%=cstr(cswday)%> 
	</a>
<tr style="height:25px">
	<td align=center style="font-size:15px;cursor:hand;font-weight:bold;">
	<img SRC='img/table.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='��榡���' style='cursor:hand;' OnClick="location.href='wk_Calendar_alldone_pr.asp?nMonth=<%=nMonth%>&nYear=<%=nYear%>'">
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
		<img SRC='img/pre_year.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='�W�@�~' style='cursor:hand;' OnClick="location.href='wk_Calendar_list_alldone_pr.asp?nMonth=<%=pre2month%>&nYear=<%=pre2year%>'">
		<img SRC='img/next_year.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='�U�@�~' style='cursor:hand;' OnClick="location.href='wk_Calendar_list_alldone_pr.asp?nMonth=<%=pre3month%>&nYear=<%=pre3year%>'">
	<td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
		<font style="vertical-align:bottom;font-size:18px;letter-spacing:2px;">�褸<%=cstr(nYear)%>�~<%=cstr(nMonth)%>��</font>
	<td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
		<img SRC='img/pre_month.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='�W�@��' style='cursor:hand;' OnClick="location.href='wk_Calendar_list_alldone_pr.asp?nMonth=<%=pre1month%>&nYear=<%=pre1year%>'">
		<img SRC='img/next_month.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='�U�@��' style='cursor:hand;' OnClick="location.href='wk_Calendar_list_alldone_pr.asp?nMonth=<%=pre4month%>&nYear=<%=pre4year%>'">
</table>
	<table border=0 bgcolor="gray" cellpadding=1 style="border-width:0px;width:700px;">
	<col style="width:50px;text-align:center;background-color:#F0FFF0;">
	<col style="width:50px;text-align:center;background-color:#F0FFF0;">
	<col style="text-align:left;background-color:#F0FFF0;">
	<tr >
		<td align=center><font style="font-size:15px;"><b>���</b></font></td>
		<td align=center><font style="font-size:15px;"><b>�P��</b></font></td>
		<td align=center><font style="font-size:15px;"><b>�ƥ�D��</b></font></td>
	<% 
		'�}�l��J��� 
		do
			'�]�w�C�� 
			select case cint(Weekday(dtDate))
			case 1
				bgc=bgc7
				strwk="��"
			case 2
				bgc=bgc1
				strwk="�@"
			case 3
				bgc=bgc1
				strwk="�G"
			case 4
				bgc=bgc1
				strwk="�T"
			case 5
				bgc=bgc1
				strwk="�|"
			case 6
				bgc=bgc1
				strwk="��"
			case 7
				bgc=bgc6
				strwk="��"
			case else 
					
			end select
				
			'����ഫ 
			if int(Day(dtDate))<10 then
				strnDay="0"&cstr(Day(dtDate))
			else
				strnDay=cstr(Day(dtDate))
			end if
			dcodeymd=dcodeym&strnDay
			dcodeymd_a=dtDate
			'����椺��g��r 
			if int(nYear)=int(Year(date()))and int(nMonth)=int(Month(date()))and int(Day(dtDate))=int(day(date())) then
				bgc="#f0ffae" 
			else			
			end if
			Response.Write "<tr bgcolor=""#ffffc0"" bgcolor="&bgc&" >"
			Response.Write "<td valign=""top"" bgcolor="&bgc&" >"& Day(dtDate)&"</td>"
			Response.Write "<td valign=""top"" bgcolor="&bgc&" >"& strwk &"</td>"
			Response.Write "<td align=left valign=""top"" bgcolor="&bgc&" style=""padding:5 5 5 5;"" >"
			'��J�ƥ��� 
				'�إ߸�Ʈw�s������	
				set rstObj1=Server.CreateObject("ADODB.Recordset")
				strSQL_show="Select * from " & tb_name & " where done_date1 = #"&dcodeymd_a&"# and wk_checker like '%"&worker&"%' order by wk_id asc"
				rstObj1.open strSQL_show,conDB,3,1
				totalput=rstObj1.recordcount
			if not rstObj1.EOF then
				rstObj1.Movefirst
				for i = 1 to totalput
					wk_id=rstObj1.fields("wk_id")
					Response.Write "<font style=""font-size:13px;"">��["& rstObj1.fields("wk_class")&"]" &rstObj1.fields("wk_item") 
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
			Response.Write "</td></tr>"
			dtDate = DateAdd("d", 1, dtDate)
				
		loop until (Month(dtDate) <> CInt(nMonth))
		'�����g���� 
			%>
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
