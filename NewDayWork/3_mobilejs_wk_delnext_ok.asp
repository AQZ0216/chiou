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
p_wk_id=Request("wk_id")         'id

%>	
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

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & p_wk_id
rstObj1.open strSQL_show,conDB,3,1
'讀取資料
p_undo_date1=rstObj1.fields("undo_date1")
p_doing_date1=rstObj1.fields("doing_date1")
p_wk_class=rstObj1.fields("wk_class")
p_wk_group=rstObj1.fields("wk_group")
p_wk_item=rstObj1.fields("wk_item")
p_wk_content=rstObj1.fields("wk_content")
p_wk_order=rstObj1.fields("wk_order")
p_wk_exe=rstObj1.fields("wk_exe")
p_wk_doer=rstObj1.fields("wk_doer")
p_wk_undoer=rstObj1.fields("wk_undoer")

'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing

'刪除資料之SQL指令字串
strSQL_del="Delete from " & tb_name & " where wk_id =" & p_wk_id
'執行SQL指令
conDB.Execute strSQL_del

'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%>
<%
'===============將要刪除之資料記錄到temp-daywork.mdb中===============
p_tmp_id=p_wk_id
p_iptok=0
' 連結Access資料庫temp-daywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="del_work_data"

'新增資料之SQL指令字串
strSQL_add="Insert into "&tb_name&" (tmp_id,"
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
strSQL_add=strSQL_add & "wk_doer,wk_undoer,ipt_ok) values ('"
strSQL_add=strSQL_add & p_tmp_id &"','"
strSQL_add=strSQL_add & p_undo_date1 &"','"
strSQL_add=strSQL_add & p_doing_date1 &"','"
strSQL_add=strSQL_add & p_wk_class &"','"
strSQL_add=strSQL_add & p_wk_group &"','"
strSQL_add=strSQL_add & p_wk_item&"','"
strSQL_add=strSQL_add & p_wk_content &"','"
strSQL_add=strSQL_add & p_wk_order &"','"
strSQL_add=strSQL_add & p_wk_exe &"','"
strSQL_add=strSQL_add & p_wk_doer&"','"
strSQL_add=strSQL_add & p_wk_undoer&"',"
strSQL_add=strSQL_add & p_iptok&")"

'執行SQL指令
conDB.Execute strSQL_add

'關閉資料庫
conDB.Close
'重設物件變數
set conDB=Nothing
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

<script language="Javascript">
	alert("資料刪除完成！！");
	location.href="wk_Calendar_all.asp";
</script>
</body>
</html>
