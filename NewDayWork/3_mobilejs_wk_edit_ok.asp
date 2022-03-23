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

p_wk_id=Request("wk_id1")         'id
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
p_doing_date1=Request("doing_date1")       '執行日期
p_wk_item=Request("wk_item")                     '主旨
p_wk_item=replace(p_wk_item,"'","’")
p_wk_content=Request("wk_content")         '內容
p_wk_content=replace(p_wk_content,"'","’")
p_all_worker=Request("all_worker")     '知會人員
p_wk_exe=request("wk_exe")       '執行人員

if  instr(1,p_all_worker,worker,1)=0 then p_all_worker=p_all_worker&","&worker
p_all_worker=replace(p_all_worker," ","",1,-1,1)

p_wk_contenta=p_wk_content

' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id="& p_wk_id &" order by wk_id desc" 
rstObj1.open strSQL_show,conDB,3,3

if rstObj1.recordcount=0 then
else	
	rstObj1.movefirst
		rstObj1.fields("doing_date1")=p_doing_date1	'執行日期
		rstObj1.fields("wk_exe")=p_wk_exe						'執行人員
		rstObj1.fields("wk_doer")=p_all_worker		'知會人員
		rstObj1.fields("wk_content")=p_wk_content		'內容
		rstObj1.fields("wk_item")=p_wk_item					'主旨
	rstObj1.UpdateBatch	
end if
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

' 連結Access資料庫temp-daywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where (tmp_id ="& p_wk_id &" and ipt_ok=0) order by wk_id desc" 
rstObj1.open strSQL_show,conDB,3,3

if rstObj1.recordcount=0 then
else	
	rstObj1.movefirst
		rstObj1.fields("doing_date1")=p_doing_date1	'執行日期
		rstObj1.fields("wk_exe")=p_wk_exe						'執行人員
		rstObj1.fields("wk_doer")=p_all_worker		'知會人員
		rstObj1.fields("wk_content")=p_wk_content		'內容
		rstObj1.fields("wk_item")=p_wk_item					'主旨
	rstObj1.UpdateBatch	
end if
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing


'關閉資料庫
conDB.Close
'重設物件變數
set conDB=Nothing
'=============================================================================
   str_url="3_mobilejs_wk_show.asp?wk_id="&p_wk_id
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
