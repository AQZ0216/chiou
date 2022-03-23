<%@ Language=VBScript CODEPAGE=950 %>
<%
function Weekday_td(wkd) 

	select case cint(wkd)
	case 1
		dcswday="日"
	case 2
		dcswday="一"
	case 3
		dcswday="二"
	case 4
		dcswday="三"
	case 5
		dcswday="四"
	case 6
		dcswday="五"
	case 7
		dcswday="六"
	end select
	Weekday_td=dcswday
end function

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
strbackURL="wk_Calendar_all_email.asp?nMonth="&nMonth&"&nYear="&nYear
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
	font-size:12.5px; 			/*字體大小*/
	background-color:'#F0FFF0'; /*背景顏色*/
	}
notetext{
	font-family:'標楷體';		/*字形*/
	font-size:12.5px; 			/*字體大小*/
	}
daytext{
	font-family:'標楷體';		/*字形*/
	font-size:12.5px; 			/*字體大小*/
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
		<b><%=worker%>未完成工作日曆表</b>
	<td align=center style="font-size:15px;cursor:hand;">
		今天是&nbsp;西元<%=cstr(Year(date()))%>年<%=cstr(Month(date()))%>月<%=cstr(Day(date()))%>日&nbsp;<%=cstr(cswday)%> 
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
		<td align=center><font color="black"><b>星期日<br>Sunday</b></font></td>
		<td align=center><font color="black"><b>星期一<br>Monday</b></font></td>
		<td align=center><font color="black"><b>星期二<br>Tuesday</b></font></td>
		<td align=center><font color="black"><b>星期三<br>Wednesday</b></font></td>
		<td align=center><font color="black"><b>星期四<br>Thursday</b></font></td>
		<td align=center><font color="black"><b>星期五<br>Friday</b></font></td>
		<td align=center><font color="black"><b>星期六<br>Saturday</b></font></td></tr>
--> 
	<tr bgcolor="#ffffc0" style="height:60px;">
		<% 
		' Add blank cells until the proper day增加空格到指定位置 
		for nDex = 1 to Weekday(dtDate) - 1
			Response.Write "<td bgcolor=""#c0c0c0"">&nbsp;</td>"
		next
		'開始填入日期 
		do
			'日期轉換 
			if int(Day(dtDate))<10 then
				strnDay="0"&cstr(Day(dtDate))
			else
				strnDay=cstr(Day(dtDate))
			end if
			dcodeymd=dcodeym&strnDay
			dcodeymd_a=dtDate
			'設定顏色 
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

			'日期格內填寫文字 
			if int(nYear)=int(Year(date()))and int(nMonth)=int(Month(date()))and int(Day(dtDate))=int(day(date())) then
				Response.Write "<td valign=""top"" bgcolor=""#f0ffae"" >"
			else			
				Response.Write "<td valign=""top"" bgcolor="&bgc&" >"
			end if
					Response.Write "<table border=0 cellpadding=0 style=""width:100%;height:100%;border-width:0px;"">"
					Response.Write "<tr style=""width:100%;height:10px;"">"
					Response.Write "<td align=center style=""background-color:#dcdcdc;font-size:13px;padding:5 5 5 5;border-width:0 0 1 0;"">"
					Response.Write month(dtDate)&" / "&Day(dtDate)&"　("&str_wkd&")"
					Response.Write "<tr><td align=left valign=top style=""padding:5 5 5 5;border-width:0px;"" >"
			'填入事件資料 
				'建立資料庫存取物件	
				set rstObj1=Server.CreateObject("ADODB.Recordset")
				strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and wk_undoer like '%"&worker&"%' order by wk_item asc , wk_id asc"
				rstObj1.open strSQL_show,conDB,3,1
				totalput=rstObj1.recordcount
			if not rstObj1.EOF then
				rstObj1.Movefirst
				for i = 1 to totalput
					wk_id=rstObj1.fields("wk_id")
					Response.Write "<font style=""font-size:13px;"">"&i&"、" &rstObj1.fields("wk_item") 
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
				
				Response.Write "</table>"
				Response.Write "</td>"
						
				if WeekDay(dtDate) = 7 then 
					Response.Write "</tr>" & vbCrLf & "<tr  bgcolor=""#ffffc0"" style=""height:60px;"">"
				end if
				dtDate = DateAdd("d", 1, dtDate)
				
		loop until (Month(dtDate) <> CInt(nMonth))
		'日期填寫完成 
			' Add blank cells to fill out the rest of the month if needed增加空格到指定位置
			if Weekday(dtDate) <> 1 then 
				for nDex = Weekday(dtDate) to 7
					Response.Write "<td bgcolor=""#C0C0C0"">&nbsp;</td>"
				next
			end if
			%>
		</tr>
	</table>
<!-- 新增備註欄位-->
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
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%> 	
</center>
</body>
</html>
