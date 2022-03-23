<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
%>
<%
'設定變數 
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
	cswday="星期日"
case 2
	cswday="星期一"
case 3
	cswday="星期二"
case 4
	cswday="星期三"
case 5
	cswday="星期四"
case 6
	cswday="星期五"
case 7
	cswday="星期六"
end select

'設定星期顏色 
bgc1="#ffffff" 	'淡黃色lightyellow 
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

'設定session("strbackURL")
strbackURL="wk_Calendar_list_alldone_pr.asp?nMonth="&nMonth&"&nYear="&nYear
session("strbackURL")=strbackURL

%>
<%
'設定上一月 
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

<!-- 開啟資料庫 -->
<%
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"
%>
<HTML>
<HEAD>
<title>樣板標題</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--讀入列印樣板檔 base_print.css  -->
	<link rel="stylesheet" type="text/css" 
		media="screen" href="./css/base_print.css" title="style_print">
	<link rel="stylesheet" type="text/css" 
		media="print" href="./css/base_print.css" title="style_print">
<style type="text/css"><!--
body{
	margin:10px 0 0 0;		/*邊緣上下左右*/
	font-family:'標楷體';		/*字形*/
	font-size:4mm; 			/*字體大小*/
	background-color:'#F0FFF0'; /*背景顏色*/
	}
notetext{
	font-family:'標楷體';		/*字形*/
	font-size:3mm; 			/*字體大小*/
	}
daytext{
	font-family:'標楷體';		/*字形*/
	font-size:3mm; 			/*字體大小*/
	}
/*連結LINK之反應*/
A:link{color:black}		/*未連結之字體顏色*/
A:visited{color:black}	/*曾經連結之字體顏色*/
A:active{color:black}	/*連結之字體顏色*/

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
		<b><%=worker%>完成工作日曆表</b>
	<td colspan=3 align=center style="font-size:15px;cursor:hand;">
	<a href="wk_calendar_list_alldone_pr.asp" style="font-size:15px;letter-spacing:1px;font-weight:bold;color:black;">
	今天是&nbsp;西元<%=cstr(Year(date()))%>年<%=cstr(Month(date()))%>月<%=cstr(Day(date()))%>日&nbsp;<%=cstr(cswday)%> 
	</a>
<tr style="height:25px">
	<td align=center style="font-size:15px;cursor:hand;font-weight:bold;">
	<img SRC='img/table.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='表格式日曆' style='cursor:hand;' OnClick="location.href='wk_Calendar_alldone_pr.asp?nMonth=<%=nMonth%>&nYear=<%=nYear%>'">
	&nbsp;
	顯示年份： 
	<select name="nYears1" onChange="mysel1" ><%
		' Note: I have set the year to be between 1999 and 2000
		for nDex = 1967 to 2067
			Response.Write "<option value=""" & nDex & """"
			if nDex = CInt(nYear) then Response.Write " selected"
			Response.Write ">" & nDex
		next %></select>
	顯示月份：
	<select name="nMonths1" onChange="mysel1" ><%
		for nDex = 1 to 12
			Response.Write "<option value=""" & nDex & """"
			if MonthName(nDex) = MonthName(nMonth) then 
				Response.Write " selected"
			end if
			Response.Write ">" & MonthName(nDex)
		next %></select>&nbsp;
	
	<td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
		<img SRC='img/pre_year.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='上一年' style='cursor:hand;' OnClick="location.href='wk_Calendar_list_alldone_pr.asp?nMonth=<%=pre2month%>&nYear=<%=pre2year%>'">
		<img SRC='img/next_year.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='下一年' style='cursor:hand;' OnClick="location.href='wk_Calendar_list_alldone_pr.asp?nMonth=<%=pre3month%>&nYear=<%=pre3year%>'">
	<td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
		<font style="vertical-align:bottom;font-size:18px;letter-spacing:2px;">西元<%=cstr(nYear)%>年<%=cstr(nMonth)%>月</font>
	<td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
		<img SRC='img/pre_month.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='上一月' style='cursor:hand;' OnClick="location.href='wk_Calendar_list_alldone_pr.asp?nMonth=<%=pre1month%>&nYear=<%=pre1year%>'">
		<img SRC='img/next_month.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='下一月' style='cursor:hand;' OnClick="location.href='wk_Calendar_list_alldone_pr.asp?nMonth=<%=pre4month%>&nYear=<%=pre4year%>'">
</table>
	<table border=0 bgcolor="gray" cellpadding=1 style="border-width:0px;width:700px;">
	<col style="width:50px;text-align:center;background-color:#F0FFF0;">
	<col style="width:50px;text-align:center;background-color:#F0FFF0;">
	<col style="text-align:left;background-color:#F0FFF0;">
	<tr >
		<td align=center><font style="font-size:15px;"><b>日期</b></font></td>
		<td align=center><font style="font-size:15px;"><b>星期</b></font></td>
		<td align=center><font style="font-size:15px;"><b>事件主旨</b></font></td>
	<% 
		'開始填入日期 
		do
			'設定顏色 
			select case cint(Weekday(dtDate))
			case 1
				bgc=bgc7
				strwk="日"
			case 2
				bgc=bgc1
				strwk="一"
			case 3
				bgc=bgc1
				strwk="二"
			case 4
				bgc=bgc1
				strwk="三"
			case 5
				bgc=bgc1
				strwk="四"
			case 6
				bgc=bgc1
				strwk="五"
			case 7
				bgc=bgc6
				strwk="六"
			case else 
					
			end select
				
			'日期轉換 
			if int(Day(dtDate))<10 then
				strnDay="0"&cstr(Day(dtDate))
			else
				strnDay=cstr(Day(dtDate))
			end if
			dcodeymd=dcodeym&strnDay
			dcodeymd_a=dtDate
			'日期格內填寫文字 
			if int(nYear)=int(Year(date()))and int(nMonth)=int(Month(date()))and int(Day(dtDate))=int(day(date())) then
				bgc="#f0ffae" 
			else			
			end if
			Response.Write "<tr bgcolor=""#ffffc0"" bgcolor="&bgc&" >"
			Response.Write "<td valign=""top"" bgcolor="&bgc&" >"& Day(dtDate)&"</td>"
			Response.Write "<td valign=""top"" bgcolor="&bgc&" >"& strwk &"</td>"
			Response.Write "<td align=left valign=""top"" bgcolor="&bgc&" style=""padding:5 5 5 5;"" >"
			'填入事件資料 
				'建立資料庫存取物件	
				set rstObj1=Server.CreateObject("ADODB.Recordset")
				strSQL_show="Select * from " & tb_name & " where done_date1 = #"&dcodeymd_a&"# and wk_checker like '%"&worker&"%' order by wk_id asc"
				rstObj1.open strSQL_show,conDB,3,1
				totalput=rstObj1.recordcount
			if not rstObj1.EOF then
				rstObj1.Movefirst
				for i = 1 to totalput
					wk_id=rstObj1.fields("wk_id")
					Response.Write "<font style=""font-size:13px;"">※["& rstObj1.fields("wk_class")&"]" &rstObj1.fields("wk_item") 
					Response.Write "</font><br>"
					rstObj1.MoveNext
					if rstObj1.EOF=true then exit for
				next
			else
			end if
				'關閉資料集
				rstObj1.Close
				'重設資料變數 
				set rstObj1=Nothing
			Response.Write "</td></tr>"
			dtDate = DateAdd("d", 1, dtDate)
				
		loop until (Month(dtDate) <> CInt(nMonth))
		'日期填寫完成 
			%>
	</table>

</form>
<%

'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%> 	
</center>
</body>
</html>
