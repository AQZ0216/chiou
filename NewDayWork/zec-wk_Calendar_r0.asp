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
<style type="text/css">
<!--
div.dayblock{
   border-collapse:collapse; 	/*邊框形式重合*/
	font-family:微軟正黑體;
	/*letter-spacing:2px;*/
	font-size:12px;
	font-weight:bold;
	color:#000000;
	/*cursor:hand;*/
	background-color:#fcfcfc;
	border: 5px solid #fcfcfc;
	/*margin:1px;*/	
   width:100%;
   /*height:150px;*/
   min-height:150px;
   /*max-height:200px;*/
   /*overflow: auto;*/
	}
-->
</style>

</HEAD>
<body class="vt-container w3-pale-blue" style="overflow:hidden;">
<center>
<form method="post" name="form1" action="">
<div class="w3-pale-blue w3-center" >
   <!--功能表-->
   <div class="w3-bar w3-blue" >
      <button onclick="url_show('zec-wk_Calendar_r0.asp?worker=<%=worker%>')" class="w3-bar-item w3-button w3-mobile" style="padding:4px;margin:0px;">回日曆表</button>
      <button onclick="url_show('zec-work_query.asp?worker=<%=worker%>')" class="w3-bar-item w3-button w3-mobile" style="padding:4px;margin:0px;">工作查詢</button>
      <button onclick="url_show('zec-work_add.asp?worker=<%=worker%>')" class="w3-bar-item w3-button w3-mobile" style="padding:4px;margin:0px;">工作新增</button>
   </div>
<!--
   <div class="w3-row w3-center w3-pale-blue ">
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-wk_Calendar_r0.asp?worker=<%=worker%>')" title="回日曆表" style="padding:4px;margin:2px;">回日曆表</button>
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-work_query.asp?worker=<%=worker%>')" title="工作查詢" style="padding:4px;margin:2px;">工作查詢</button>
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-work_add.asp?worker=<%=worker%>')" title="工作新增" style="padding:4px;margin:2px;">工作新增</button>
   </div> 
-->       
   <!--內容-->
<%
p_year=request("p_year")'年
p_month=request("p_month")'月
p_day=request("p_day")'日
'p_week=request("p_week")'週數
if p_year="" or isnull(p_year) then p_year=year(date())
if p_month="" or isnull(p_month) then p_month=month(date())
if p_day="" or isnull(p_day) then p_day=day(date())

pn_date=dateserial(p_year,p_month,p_day)'今日

p_year=year(pn_date)'年
p_month=month(pn_date)'月
p_week=DatePart("ww",pn_date)'週數
p_date=pn_date'今日
p_showtype="month"   'month、week、date
p_mfweek=DatePart("ww",dateserial(p_year,p_month,1))'本月1日週數
p_mfweekday=Weekday(dateserial(p_year,p_month,1))'本月第一日星期

'已知年分+週數，查詢週數第一天(周日為第一天)
function findwk01(pp_yy,pp_wks)
   pp_wk01=DatePart("w",dateserial(pp_yy,1,1))'元旦星期幾
   pp_wk01dayno=(7-pp_wk01)+1'第一週天數
   pp_dayno=(pp_wks-2)*7+pp_wk01dayno   '日期數
   findwk01=DateAdd("d", pp_dayno, dateserial(pp_yy,1,1)) 
end function
p_firstday=dateserial(p_year,p_month,1-(p_mfweekday)+1)'顯示週數第一天
%>
      <div class="w3-row ">
         <div class="w3-col l74 w3-center w3-grey w3-border w3-border-black ">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【<<】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【<】</button>
            <%=p_year%>年<%=p_month%>月
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【今日=<%=pn_date%>】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【>】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【>>】</button>
         </div>
         <div class="w3-col l73 w3-center w3-grey w3-border w3-border-black ">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【月】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【週】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【日】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【表】</button>
         </div>
      </div>
   <div class="w3-row w3-center w3-pale-red" style="max-height:460px;overflow:scroll;">
<!--
      <div class="w3-row " style="overflow:auto;">
         <div class="w3-col l71 w3-center w3-pale-red w3-border w3-border-black " style="padding:4px;margin:0px;">星期日</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">星期一</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">星期二</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">星期四</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">星期四</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">星期五</div>
         <div class="w3-col l71 w3-center w3-pale-red w3-border w3-border-black " style="padding:4px;margin:0px;">星期六</div>
      </div> -->
<%
for wkno=1 to 6
   pn_wkno=p_mfweek+wkno-1
   pw_day01=findwk01(p_year,pn_wkno)'本周第一天
  
   if month(pw_day01)=p_month or month(pw_day01+6)=p_month then
      if pn_wkno=p_week then 
         div_background_c="#fffed9"
         div_border_c="#fffed9"
      else
         div_background_c="#fcfcfc"
         div_border_c="#fcfcfc"
      end if

   for dn=1 to 7
      Select Case dn
         Case 1    
            str_wk="日"
            div_background_c="#ffdddd"    'w3-pale-red
            div_border_c="#ffdddd"        'w3-pale-red
         Case 2    
            str_wk="一"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
         Case 3    
            str_wk="二"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
        Case 4    
            str_wk="三"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
         Case 5    
            str_wk="四"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
         Case 6    
            str_wk="五"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
        Case 7    
            str_wk="六"
            div_background_c="#ffdddd"    'w3-pale-red
            div_border_c="#ffdddd"        'w3-pale-red
         Case Else     
         
      End Select   

      pndate=pw_day01+dn-1'日期
      
      if month(pndate)=p_month then
         '本月日期
         'div_background_c="#fcfcfc"
         'div_border_c="#fcfcfc"
      else
         '非本月日期
         div_background_c="#dbdbdb"
         div_border_c="#dbdbdb"
      end if
      
      if pndate=pn_date then 
         div_background_c="#fffed9" 
         div_border_c="#fffed9"
      end if
%>   
      <div class="w3-col l71 w3-center w3-border w3-border-black dayblock" style="background-color:<%=div_background_c%>;border-color:<%=div_border_c%>;">
         <div class="w3-container" style="overflow:auto;">
            <div class="w3-col s6">
               <%=pw_day01+dn-1%>(<%=str_wk%>)
            </div>
            <div class="w3-col s6">
               <button class="w3-button w3-grey w3-medium " style="padding:0px;margin:0px;">【新增】</button>
            </div>
            <div class="w3-row w3-left " style="overflow: auto;">
               <%
               str_hdman=hd_man(pndate)
               response.write str_hdman
               %>
            </div> 
          </div>          
      </div>
<% 
   next 
   end if
next
%>
   </div>

</div>

</form>
<script language="JavaScript">
    function url_new(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        //window.location.href = pp_url; //原頁面更新
        window.open(pp_url) ; //開啟新頁面
        return true;
    }   
    function url_show(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        window.location.href = pp_url; //原頁面更新
        //window.open(pp_url) ; //開啟新頁面
        return true;
    }   
    function zms_show(pp_url){
        var iframe1=document.getElementById("ifrm_milestone");
        iframe1.src=pp_url;
        return true;
    }    
    function zlb_show(pp_url){
        var iframe1=document.getElementById("ifrm_logbook");
        iframe1.src=pp_url;
        return true;
    }
    function zfi_show(pp_url){
        var iframe1=document.getElementById("ifrm_finance");
        iframe1.src=pp_url;
        return true;
    }
</script>
</center>
</body>
</html>
