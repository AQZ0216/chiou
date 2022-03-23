<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<%
'ead會議公告
	'讀取人員姓名
	worker = Session("worker")
'輸入月份
j_date=request("p_ymd")

'判斷是否是重複性公告
end_date=dateadd("m",1,j_date)
 str_date=""
 pre_date=dateadd("d",-1,j_date)   '開始日期
 ntk=1
'每月周一至週五
      do
         next_date=dateadd("d",ntk,pre_date)
         if next_date >= cdate(end_date) then
            check_s=true 
         else
            if Weekday(next_date) > 1 and Weekday(next_date) < 7 then
               str_date=str_date&","&next_date
            end if
            ntk=ntk+1
         end if
      loop until check_s=true

'日期陣列
if left(str_date,1)="," then str_date=right(str_date,len(str_date)-1)
date_arr=Split(str_date, ",", -1, 1)
date_num=ubound(date_arr)+1
%>	
<%
'所有人員字串
     str_allworker=worker_a(0)
	for i=2 to worker_no
		str_allworker=str_allworker&","& worker_a(i-1)
	next
%>
<html>
<head>
<title>資料完整新增</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body {  scrollbar-3dlight-color:#ffffff;
        scrollbar-arrow-color:#CCCCCC;
        scrollbar-base-color:#666633;
        scrollbar-darkshadow-color:#e6e6cc;
        scrollbar-face-color:#666666;
        scrollbar-highlight-color:#ffffff;
        scrollbar-shadow-color:#e6e6cc;
        scrollbar-track-color:#ffffff;
        margin:2mm 0mm 0mm 0mm;		/*邊緣上下左右*/
		font-family:'標楷體';		/*字形*/
		font-size:4.5mm; 			/*字體大小*/
		background-color:'#F0FFF0';
     }
td{
   margin:0 0 0 0;      /*邊緣上下左右*/
   border-color:'#808080'; /*表格外框顏色*/ 
   border-style:solid;     /*表格外框線型*/
   border-width:1px;    /*表格外框厚度*/  
   vertical-align:middle;  /*字體垂直對齊方式*/
   font-size:4.5mm;
   }
table{
   margin:0 0 0 0;      /*邊緣上下左右*/
   border-collapse:collapse;  /*邊框形式重合*/
   }
--></style>
</head>
<body>
<center>

<!-- 開啟資料庫 -->
<%
'
 pct_date=date()   '公告日期
 pdt_date=date()   '執行日期
'公告日期
p_undo_date1=pct_date
'執行日期
'p_doing_date1= pdt_date
'工作分類
p_wk_class=""
'工作群組
p_wk_group="一般工作"
'主旨
'p_wk_item="08:20-09:00 EAD會議"
'p_wk_item="08:45-09:15 EAD會議"	'2014/11/24 開始
p_wk_item="08:30-09:0 E0AD會議"	'2014/12/16 開始
'執行內容
'p_wk_content="每天08:20-09:00 EAD會議"&chr(13)&"/會議期間，請勿打擾"
'p_wk_content="每天08:45-09:15 EAD會議"&chr(13)&"會議期間，請勿打擾"		'2014/11/24 開始
p_wk_content="每天08:30-09:00 EAD會議"&chr(13)&"會議期間，請勿打擾"		'2014/12/16 開始
'公告者
p_wk_order=worker

p_wk_exe="郭總,國賢,國哲"
'知會人員
p_all_worker=str_allworker     '知會人員

'response.write "公告日期p_undo_date1="&p_undo_date1&"。<br>"
'response.write "執行日期p_doing_date1="&str_date&"。<br>"
'response.write "工作群組p_wk_group="&p_wk_group&"。<br>"
'response.write "工作主旨p_wk_item="&p_wk_item&"。<br>"
'response.write "執行內容p_wk_content="&p_wk_content&"。<br>"
'response.write "公告者p_wk_order="&p_wk_order&"。<br>"
'response.write "知會人員p_all_worker="&p_all_worker&"。<br>"
'response.end
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"

for zki=1 to date_num


'新增資料之SQL指令字串
strSQL_add="Insert into "&tb_name&" ("
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
strSQL_add=strSQL_add & "wk_doer,wk_undoer) values ('"

strSQL_add=strSQL_add & p_undo_date1 &"','"
strSQL_add=strSQL_add & date_arr(zki-1) &"','"
strSQL_add=strSQL_add & p_wk_class &"','"
strSQL_add=strSQL_add & p_wk_group &"','"

strSQL_add=strSQL_add & p_wk_item&"','"
strSQL_add=strSQL_add & p_wk_content &"','"
strSQL_add=strSQL_add & p_wk_order &"','"
strSQL_add=strSQL_add & p_wk_exe &"','"
strSQL_add=strSQL_add & p_all_worker&"','"
strSQL_add=strSQL_add & p_all_worker&"')"

'執行SQL指令
conDB.Execute strSQL_add

next

'關閉資料庫
conDB.Close
'重設物件變數
set conDB=Nothing
   str_url="wk_calendar_all.asp?nMonth="& month(j_date) &"&nYear="& year(j_date)
   response.redirect(str_url)
%>
</body>
</html>
