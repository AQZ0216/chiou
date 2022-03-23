<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
'判斷是否輸入工作分類 
keyword=request("wk_class")
'if keyword="" then 
	'response.redirect("wk_add.asp")
'else
'end if


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
p_undo_date1=Request("undo_date1")         '公告日期
p_doing_date1=Request("doing_date1")       '執行日期
p_wk_class=Request("wk_class")                   '工作分類
p_wk_group=Request("wk_group")                '工作群組
p_wk_item=Request("wk_item")                     '主旨
p_wk_item=replace(p_wk_item,"'","’")
p_wk_content=Request("wk_content")         '內容
p_wk_content=replace(p_wk_content,"'","’")
p_wk_order=Request("wk_order")                 '公告者
p_all_worker=Request("all_worker")     '知會人員
	p_wk_exe=request("wk_exe")       '執行人員
if  instr(1,p_all_worker,worker,1)=0 then p_all_worker=p_all_worker&","&worker
p_all_worker=replace(p_all_worker," ","",1,-1,1)

'===========判斷資料是否填寫完整=================
str_error=""
if  p_doing_date1="" or isnull(p_doing_date1) or not(isdate(p_doing_date1)) then str_error=str_error&"[執行日期]錯誤。"
if  p_wk_item="" or isnull(p_wk_item) then str_error=str_error&"[主旨]空白。"
'if  p_wk_content="" or isnull(p_wk_content) then str_error=str_error&"[內容]空白。"
if  p_all_worker="" or isnull(p_all_worker) then str_error=str_error&"[知會人員]空白。"
if  p_wk_order="" or isnull(p_wk_order) then response.redirect("./firstpage.asp")
if  p_wk_exe="" or isnull(p_wk_exe) then str_error=str_error&"[執行人員]空白。"

if not(str_error="") then
   nexturl="3_mobilejs_wk_add.asp?ermsg="&str_error
   response.redirect(nexturl)
end if
'===========判斷資料是否填寫完整=================
p_wk_pjn=Request("wk_pjn")     '專案名稱

if p_wk_pjn="0" or isnull(p_wk_pjn) then
      p_pj_id=null
      p_pj_02=null
else
      a_wk_pjn=split(p_wk_pjn,"，",-1,1)
      p_pj_id=a_wk_pjn(0)
      p_pj_02=a_wk_pjn(1)
end if

if  instr(1,p_all_worker,worker,1)=0 then p_all_worker=p_all_worker&","&worker

p_wk_contenta=p_wk_content
'response.write "p_undo_date1=" & p_undo_date1 & "<br>"
'response.write "p_doing_date1=" & p_doing_date1 & "<br>"
'response.write "p_wk_class=" & p_wk_class & "<br>"
'response.write "p_wk_exe=" & p_wk_exe & "<br>"
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

'新增資料之SQL指令字串
strSQL_add="Insert into "&tb_name&" ("
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
if p_wk_pjn="0" or isnull(p_wk_pjn) then
else
   strSQL_add=strSQL_add & "pj_id,pj_02,"
end if
strSQL_add=strSQL_add & "wk_doer,wk_undoer) values ('"

strSQL_add=strSQL_add & p_undo_date1 &"','"
strSQL_add=strSQL_add & p_doing_date1 &"','"
strSQL_add=strSQL_add & p_wk_class &"','"
strSQL_add=strSQL_add & p_wk_group &"','"

strSQL_add=strSQL_add & p_wk_item&"','"
strSQL_add=strSQL_add & p_wk_contenta &"','"
strSQL_add=strSQL_add & p_wk_order &"','"
strSQL_add=strSQL_add & p_wk_exe &"','"
if p_wk_pjn="0" or isnull(p_wk_pjn) then
else
   strSQL_add=strSQL_add & p_pj_id &"','"
   strSQL_add=strSQL_add & p_pj_02 &"','"
end if

strSQL_add=strSQL_add & p_all_worker&"','"
strSQL_add=strSQL_add & p_all_worker&"')"

'執行SQL指令
conDB.Execute strSQL_add

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " order by wk_id desc" 
rstObj1.open strSQL_show,conDB,1,1
rstObj1.movefirst
	p_tmp_id=rstObj1.fields("wk_id")
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing

'關閉資料庫
conDB.Close
'重設物件變數
set conDB=Nothing
'=============================================================================
'暫存檔案
p_iptok=0
' 連結Access資料庫temp-daywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"

'新增資料之SQL指令字串
strSQL_add="Insert into "&tb_name&" (tmp_id,"
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
if p_wk_pjn="0" or isnull(p_wk_pjn) then
else
   strSQL_add=strSQL_add & "pj_id,pj_02,"
end if
strSQL_add=strSQL_add & "wk_doer,wk_undoer,ipt_ok) values ('"

strSQL_add=strSQL_add & p_tmp_id &"','"
strSQL_add=strSQL_add & p_undo_date1 &"','"
strSQL_add=strSQL_add & p_doing_date1 &"','"
strSQL_add=strSQL_add & p_wk_class &"','"
strSQL_add=strSQL_add & p_wk_group &"','"

strSQL_add=strSQL_add & p_wk_item&"','"
strSQL_add=strSQL_add & p_wk_contenta &"','"
strSQL_add=strSQL_add & p_wk_order &"','"
strSQL_add=strSQL_add & p_wk_exe &"','"
if p_wk_pjn="0" or isnull(p_wk_pjn) then
else
   strSQL_add=strSQL_add & p_pj_id &"','"
   strSQL_add=strSQL_add & p_pj_02 &"','"
end if

strSQL_add=strSQL_add & p_all_worker&"','"
strSQL_add=strSQL_add & p_all_worker&"',"
strSQL_add=strSQL_add & p_iptok&")"

'執行SQL指令
conDB.Execute strSQL_add

'關閉資料庫
conDB.Close
'重設物件變數
set conDB=Nothing
'=============================================================================
   str_url="3_mobilejs_wk_show.asp?wk_id="&p_tmp_id
'   response.redirect(str_url)
'   str_url="wk_calendar_all.asp"
   response.redirect(str_url)

%>
<!-- <script language="Javascript">
	alert("資料新增完成！！");
	location.href="wk_calendar_all.asp";
</script> -->

</body>
</html>
