<%@ Language=VBScript CODEPAGE=950 %>
<%
   '讀取人員姓名
   worker = request("worker")
%>
<%
'休假資料
function hd_man(p_hdate)
   pstr_hdman =""
    ' 連結Access資料庫holiday.mdb
    DBpath_fh=Server.MapPath("../holiday/database/holiday.mdb")
    strCon_fh="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fh
    '建立資料庫連結物件
    set conDB_fh= Server.CreateObject("ADODB.Connection")
    '連結資料庫	
    conDB_fh.Open strCon_fh
    '開啟資料表名稱
    tb_name_fh="休假明細"
	'建立資料庫存取物件
	set rstObj1_fh=Server.CreateObject("ADODB.Recordset")
	strSQL_show_fh="Select * from " & tb_name_fh & " where 休假日 = #"& p_hdate &"# order by 假別id asc "
	rstObj1_fh.open strSQL_show_fh,conDB_fh,3,1
	totalput_fh=rstObj1_fh.recordcount
if not rstObj1_fh.EOF then
	rstObj1_fh.Movefirst
	for i = 1 to totalput_fh
		hd_id=rstObj1_fh.fields("hd_id")
		icon_id=rstObj1_fh.fields("假別id")
		hd_hrs=rstObj1_fh.fields("休假時數")
		hd_check=rstObj1_fh.fields("確認")
		hd_man=rstObj1_fh.fields("員工姓名")'員工姓名
		hd_img=left(rstObj1_fh.fields("假別名稱"),1)
		hd_cname=right(rstObj1_fh.fields("假別名稱"),len(rstObj1_fh.fields("假別名稱"))-1)
		'決定假別顏色
		select case icon_id
		   Case 1  f_color = "#000000"    '○：公休。
		   Case 2  f_color = "#000000"    '▲：事假。
		   Case 3  f_color = "#000000"    '⊕：病假。
		   Case 4  f_color = "#000000"    '㊣：公假。
		   Case 5  f_color = "#000000"    '◆：喪假。
		   Case 6  f_color = "#000000"    '△：年假。
		   Case 7  f_color = "#000000"    '■：特休。
		   Case 8  f_color = "#000000"    '★：產假。
		   Case 9  f_color = "#000000"    '◎：婚假。
		   Case 15  f_color = "#000000"   '※：未打卡。
		   Case 16  f_color = "#000000"   '▽：排班。
		   Case 17  f_color = "#000000"    '＠：產檢假。
		   Case 18  f_color = "#000000"    '＠：陪產假。
		   Case 19  f_color = "#000000"    '♀：育嬰假。
		   Case Else   f_color = "#000000"
		End Select
		if icon_id=1 or icon_id=15 then
		    if icon_id=15 then
		       pstr_hdman = pstr_hdman & "<font style='font-size:15px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man & "</font><br>"
   	           end if
		else
		  pstr_hdman = pstr_hdman & "<font style='font-size:15px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man &"("& hd_hrs&")&nbsp;</font><br>"
		end if
		    'Response.Write "</font><br>"
		rstObj1_fh.MoveNext
		if rstObj1_fh.EOF=true then exit for
	next
else
end if
	'關閉資料集
	rstObj1_fh.Close
	'重設資料變數 
	set rstObj1_fh=Nothing
    '關閉資料庫
    conDB_fh.Close
    '重設物件變數 
    set conDB_fh=Nothing
  hd_man=pstr_hdman
end function
%>
<%
'查詢是否有附件
Function exist_attach(pwk_id)
      ' 連結Access資料庫daywork.mdb
      DBpath_fe=Server.MapPath("./database/attach_file.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '建立資料庫連結物件
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '連結資料庫	
      conDB_fe.Open strCon_fe
      '開啟資料表名稱
      tb_name_fe="file_data"
      '建立資料庫存取物件	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where del_ok = false and wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '關閉資料集
      rstObj1_fe.Close
      '重設資料變數
      set rstObj1_fe=Nothing
      '關閉資料庫 
      conDB_fe.Close
      '重設物件變數
      set conDB_fe=Nothing
      exist_attach=totalput_fe
End Function

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
bgc1="#F0FFF0"    '淡黃色lightyellow 
bgc6="#F0FFF0" 'lightskyblue
bgc7="#F0FFF0" 'lightgreen

' Set the date to the first of the current month
dtDate = DateSerial(nYear, nMonth, 1)


if int(nMonth)<10 then
   strnMonth="0"&cstr(nMonth)
else
   strnMonth=cstr(nMonth)
end if
dcodeym=cstr(nYear)&strnMonth

'設定session("strbackURL")
strbackURL="wk_Calendar_all.asp?nMonth="&nMonth&"&nYear="&nYear
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
<%
pstart=dateserial(nYear,nMonth,1)
pend=dateadd("m",1,pstart)
'查詢本月是否建立EAD會議         p_wk_item="08:20-09:00 EAD會議"
function find_ead(pstart,pend)
      ' 連結Access資料庫daywork.mdb
      DBpath_ead=Server.MapPath("./database/daywork.mdb")
      strCon_ead="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_ead
      '建立資料庫連結物件
      set conDB_ead= Server.CreateObject("ADODB.Connection")
      '連結資料庫   
      conDB_ead.Open strCon_ead
      '開啟資料表名稱
      tb_name_ead="work_data"
      '建立資料庫存取物件	
      set rstObj1_ead=Server.CreateObject("ADODB.Recordset")
      strSQL_show_ead="Select * from " & tb_name_ead & " where wk_item like '08:20-09:00 EAD會議' and wk_order like '美慧' and doing_date1 >= #"& pstart &"# and doing_date1 < #"& pend &"# order by doing_date1 asc"
      rstObj1_ead.open strSQL_show_ead,conDB_ead,3,1
      totalput_ead=rstObj1_ead.recordcount
      '關閉資料集
      rstObj1_ead.Close
      '重設資料變數 
      set rstObj1_ead=Nothing
      '關閉資料庫 
      conDB_ead.Close
      '重設物件變數 
      set conDB_ead=Nothing
      find_ead=totalput_ead
end function
pchk_ead=find_ead(pstart,pend)
%>

<HTML>
<HEAD>
<title>樣板標題</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="./css/w3-cht.css">
<style type="text/css"><!--

--></style>

</HEAD>
<body class="vt-container" >
<center>
<form method="post" name="form1" action="">


<table border=0 bgcolor="gray" cellpadding=1 style="width:100%;border-width:0px;width:970px;">
<col style="width:56%;background-color:#F0FFF0;">
<col style="width:11%;background-color:#F0FFF0;">
<col style="width:22%;background-color:#F0FFF0;">
<col style="width:11%;background-color:#F0FFF0;">
<tr style="height:25px">
   <td align=center style="font-size:18px;letter-spacing:1px;" >
       <img SRC='img/cdfundation.ico' align=middle WIDTH='24' HEIGHT='24' BORDER='0' ALT='基金會活動紀錄' style='cursor:hand;' OnClick="window.open('z_cd_recordlist.asp')">
   <%
if (worker="美慧") and pchk_ead=0  then
   p_ymd=dateserial(nYear,nMonth,1)
%>
<a href="./wk_add_EAD_ok.asp?p_ymd=<%=p_ymd%>" target="_self" title="建立非假日之EAD會議。">EAD會議</a>
<%
end if
   %>
      <b><%=worker%>未完成工作日曆表</b>
      <a href="wk_calendar_all_pr.asp?nMonth=<%=nMonth%>&nYear=<%=nYear%>" target="_blank" style="font-size:3.5mm;letter-spacing:1px;color:red;">[友善列印]</a>
      <a href="wk_calendar_all_pru.asp?nMonth=<%=nMonth%>&nYear=<%=nYear%>" target="_blank" style="font-size:3.5mm;letter-spacing:1px;color:red;">上</a>
      <a href="wk_calendar_all_prd.asp?nMonth=<%=nMonth%>&nYear=<%=nYear%>" target="_blank" style="font-size:3.5mm;letter-spacing:1px;color:red;">下</a>
      <a href="wk_calendar_all_email.asp?nMonth=<%=nMonth%>&nYear=<%=nYear%>" target="_blank" style="font-size:3.5mm;letter-spacing:1px;color:red;">[email]</a>
   <td colspan=3 align=center style="font-size:15px;cursor:hand;">
   <a href="wk_calendar_all.asp" style="font-size:15px;letter-spacing:1px;font-weight:bold;color:black;">
   今天是&nbsp;西元<%=cstr(Year(date()))%>年<%=cstr(Month(date()))%>月<%=cstr(Day(date()))%>日&nbsp;<%=cstr(cswday)%> 
   </a>
       <img SRC='img/gcalendar.png' align=middle WIDTH='20' HEIGHT='20' BORDER='0' ALT='匯出工作項目為Google日曆csv檔' style='cursor:hand;' OnClick="window.open('zwk2google_query.asp')">
<tr style="height:25px">
   <td align=center style="font-size:15px;cursor:hand;font-weight:bold;">
   <img SRC='img/list.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='條列式日曆' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_list_all.asp?nMonth=<%=nMonth%>&nYear=<%=nYear%>'">
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
      <img SRC='img/pre_year.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='上一年' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_all.asp?nMonth=<%=pre2month%>&nYear=<%=pre2year%>'">
      <img SRC='img/next_year.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='下一年' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_all.asp?nMonth=<%=pre3month%>&nYear=<%=pre3year%>'">
   <td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
      <font style="vertical-align:bottom;font-size:18px;letter-spacing:2px;">西元<%=cstr(nYear)%>年<%=cstr(nMonth)%>月</font>
   <td align=center style="font-size:19px;cursor:hand;font-weight:bold;">
      <img SRC='img/pre_month.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='上一月' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_all.asp?nMonth=<%=pre1month%>&nYear=<%=pre1year%>'">
      <img SRC='img/next_month.bmp' align=middle WIDTH='30' HEIGHT='20' BORDER='0' ALT='下一月' style='cursor:hand;' OnClick="parent.content.location.href='wk_Calendar_all.asp?nMonth=<%=pre4month%>&nYear=<%=pre4year%>'">
</table>
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
   <table border=0 bgcolor="gray" cellpadding=1 style="width:100%;border-width:0px;width:970px;">
   <col style="width:14%;background-color:<%=bgc7%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc6%>;">
   <tr style="height:45px;">
      <td align=center><font color="black"><b>星期日<br>Sunday</b></font></td>
      <td align=center><font color="black"><b>星期一<br>Monday</b></font></td>
      <td align=center><font color="black"><b>星期二<br>Tuesday</b></font></td>
      <td align=center><font color="black"><b>星期三<br>Wednesday</b></font></td>
      <td align=center><font color="black"><b>星期四<br>Thursday</b></font></td>
      <td align=center><font color="black"><b>星期五<br>Friday</b></font></td>
      <td align=center><font color="black"><b>星期六<br>Saturday</b></font></td></tr>
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

         '日期格內填寫文字 
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
               Response.Write "<font style=""font-size:13px;""><a href='wk_add.asp?datecode="&dtDate&"'>新增"
               Response.Write "</a></font>"
               'Response.Write "<font style=""text-size:3mm;""><b>&nbsp;"
               'Response.Write "</b></font>"
               Response.Write "<tr><td colspan=2 align=left valign=top>"
         '填入事件資料 
            'Response.Write "<font size=""-1""><b>" &dcodeymd&"</b></font><br>"
            'Response.Write "<font size=""-1""><b>" &totalput&"</b></font><br>"
            '建立資料庫存取物件  
            set rstObj1=Server.CreateObject("ADODB.Recordset")

'            strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and wk_undoer like '%"&worker&"%' order by wk_item asc , wk_id asc"
if dcodeymd_a>= date() then
            strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and (wk_undoer like '%"&worker&"%' or wk_exe like '%"&worker&"%' or wk_att like '%"&worker&"%') order by wk_item asc , wk_id asc"
else
'            strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and (wk_undoer like '%"&worker&"%' or wk_exe like '%"&worker&"%' or wk_att like '%"&worker&"%') order by wk_item asc , wk_id asc"
            strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and wk_undoer like '%"&worker&"%' order by wk_item asc , wk_id asc"
'response.write strSQL_show
'response.end
end if
            rstObj1.open strSQL_show,conDB,3,1
            totalput=rstObj1.recordcount
         if not rstObj1.EOF then
            rstObj1.Movefirst
            for i = 1 to totalput
            	 wk_headline=rstObj1.fields("headline")  '跑馬燈
               wk_id=rstObj1.fields("wk_id")
               '檢查是否有附件 exist_attach(wk_id)
               attach_no=exist_attach(wk_id)
               if attach_no=0 then
                  str_colors="color:#000000;"
               else
                  str_colors="color:#0000FF;"
               end if
               if rstObj1.fields("wk_password")="" or isnull(rstObj1.fields("wk_password")) then
               else
                  str_colors="color:#0000FF;"
               end if
								p_nexe=rstObj1.fields("wk_exe")	'執行人員
								if Instr(1, p_nexe, worker, 1)>0 or Instr(1, p_nexe, "全體", 1)>0 then
									str_bgc="background-color:#99FF99;"	
								else
									str_bgc=""
								end if
               Response.Write "<font style='font-size:13px;"& str_bgc & str_colors &"' >" & i & "、<a href='wk_show.asp?wk_id="&wk_id&"' style='letter-spacing:1.5pt;font-size:11pt;"& str_colors &"' >" & replace (rstObj1.fields("wk_item"),"平菁雲","<font color=fuchsia >平菁雲</font>")
               Response.Write "</a></font>"
if wk_headline>5 then response.write "<img src='./img/gnome_chess.png' title='跑馬燈訊息' width=19 style='vertical-align:top;'>"                
               Response.Write "<br>"
               rstObj1.MoveNext
               if rstObj1.EOF=true then exit for
            next
         else
         end if
            '關閉資料集
            rstObj1.Close
            '重設資料變數 
            set rstObj1=Nothing
             str_hdman=hd_man(dcodeymd_a)     '休假人員資料
            response.write str_hdman
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
      <tr>
      <td colspan=7>
<b>
<font color="#000000">○：公休。</font>
<font color="#000000">▲：事假。</font>
<font color="#000000">⊕：病假。</font>
<font color="#000000">㊣：公假。</font>
<font color="#000000">◆：喪假。</font>
<font color="#000000">△：年假。</font>
<font color="#000000">■：特休。</font>
<font color="#000000">★：產假。</font><BR>
<font color="#000000">＠：產檢假。</font>
<font color="#000000">＠：陪產假。</font>
<font color="#000000">◎：婚假。</font>
<font color="#000000">♀：育嬰假。</font>
<font color="#000000">※：未打卡。</font>
<font color="#000000">▽：排班。</font>
<font color="#000000">()：請假時數。</font>
</b>
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
