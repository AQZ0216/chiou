<%@ Language=VBScript CODEPAGE=950 %>
<%
'讀取temp-daywork.mdb資料  '==================================
'pa_undo_date1=Request("undo_date1")         '公告日期
'pa_doing_date1=Request("doing_date1")       '執行日期
'pa_wk_class=Request("wk_class")                   '工作分類
'pa_wk_group=Request("wk_group")                '工作群組
'pa_wk_item=Request("wk_item")                     '主旨
'pa_wk_content=Request("wk_content")         '內容
'pa_wk_order=Request("wk_order")                 '公告者
'pa_all_worker=Request("all_worker")     '知會人員
dim pa_undo_date1()   '公告日期
dim pa_doing_date1()      '執行日期
dim pa_wk_class()                  '工作分類
dim pa_wk_group()               '工作群組
dim pa_wk_item()                    '主旨
dim pa_wk_content()         '內容
dim pa_wk_order()                '公告者
dim pa_wk_doer()                '工作人員知會人員
dim pa_wk_undoer()                '未完成工作人員
dim pa_wk_exe()                '執行人員

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
	strSQL_show="Select * from " & tb_name & " where ipt_ok = 0"
	rstObj1.open strSQL_show,conDB,3,3
	'計算資料總數	
	totalput=rstObj1.recordcount
	if totalput=0 then
	else
redim pa_undo_date1(totalput)   '公告日期
redim pa_doing_date1(totalput)      '執行日期
redim pa_wk_class(totalput)                  '工作分類
redim pa_wk_group(totalput)               '工作群組
redim pa_wk_item(totalput)                    '主旨
redim pa_wk_content(totalput)         '內容
redim pa_wk_order(totalput)                '公告者
redim pa_wk_doer(totalput)                '工作人員知會人員
redim pa_wk_undoer(totalput)                '未完成工作人員
redim pa_wk_exe(totalput)                '執行人員
		'列出資料項目
		rstobj1.MoveFirst
		for j=1 to totalput
			pa_undo_date1(j-1)=rstObj1.fields("undo_date1")         '公告日期
			pa_doing_date1(j-1)=rstObj1.fields("doing_date1")         '執行日期
			pa_wk_class(j-1)=rstObj1.fields("wk_class")         '工作分類
			pa_wk_group(j-1)=rstObj1.fields("wk_group")         '工作群組
			pa_wk_item(j-1)=rstObj1.fields("wk_item")         '主旨
			pa_wk_content(j-1)=rstObj1.fields("wk_content")         '內容
			pa_wk_order(j-1)=rstObj1.fields("wk_order")         '公告者
			pa_wk_doer(j-1)=rstObj1.fields("wk_doer")         '工作人員知會人員
			pa_wk_undoer(j-1)=rstObj1.fields("wk_undoer")         '未完成工作人員
			pa_wk_exe(j-1)=rstObj1.fields("wk_exe")         '執行人員
			'移到下一筆記錄
			rstObj1.MoveNext
			if rstObj1.EOF=True then exit for
		next
	end if
	'關閉資料集
	rstObj1.Close
	'重設資料變數 
	set rstObj1=Nothing
	'關閉資料庫
	conDB.Close
	'重設物件變數
	set conDB=Nothing
'==================================

if totalput>0 then              '==================================
   ' 連結Access資料庫daywork.mdb
   DBpath=Server.MapPath("./database/daywork.mdb")
   strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
   '建立資料庫連結物件
   set conDB= Server.CreateObject("ADODB.Connection")
   '連結資料庫	
   conDB.Open strCon
   '開啟資料表名稱
   tb_name="work_data"

      for kj=1 to totalput          '==================================
         p_undo_date1=pa_undo_date1(kj-1)
         p_doing_date1=pa_doing_date1(kj-1)
         p_wk_class=pa_wk_class(kj-1)
         p_wk_group=pa_wk_group(kj-1)
         p_wk_item=pa_wk_item(kj-1)
         'p_wk_content=pa_wk_content(kj-1)
         p_wk_content=pa_wk_content(kj-1) & chr(13) & date()& "手機新增。" 
         p_wk_order=pa_wk_order(kj-1)
         p_wk_doer=pa_wk_doer(kj-1)
         p_wk_undoer=pa_wk_undoer(kj-1)
         p_wk_exe=pa_wk_exe(kj-1)
         '新增資料之SQL指令字串
         strSQL_add="Insert into "&tb_name&" ("
         strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
         strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
         strSQL_add=strSQL_add & "wk_doer,wk_undoer) values ('"
         strSQL_add=strSQL_add & p_undo_date1 &"','"
         strSQL_add=strSQL_add & p_doing_date1 &"','"
         strSQL_add=strSQL_add & p_wk_class &"','"
         strSQL_add=strSQL_add & p_wk_group &"','"
         strSQL_add=strSQL_add & p_wk_item &"','"
         strSQL_add=strSQL_add & p_wk_content &"','"
         strSQL_add=strSQL_add & p_wk_order &"','"
         strSQL_add=strSQL_add & p_wk_exe &"','"
         strSQL_add=strSQL_add & p_wk_doer&"','"
         strSQL_add=strSQL_add & p_wk_undoer&"')"
         '執行SQL指令
         conDB.Execute strSQL_add
      next                        '==================================
   '關閉資料庫
   conDB.Close
   '重設物件變數
   set conDB=Nothing
end if                            '==================================

%>
<%
'=========刪除已同步之資料==========================
'讀取temp-daywork.mdb資料  '==================================
dim pa_delwkid()   '刪除之工作id

	' 連結Access資料庫temp-daywork.mdb
	DBpath=Server.MapPath("./database/temp-daywork.mdb")
	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'建立資料庫連結物件
	set conDB= Server.CreateObject("ADODB.Connection")
	'連結資料庫	
	conDB.Open strCon
	'開啟資料表名稱
	tb_name="del_work_data"
	'建立資料庫存取物件	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " where ipt_ok = 0"
	rstObj1.open strSQL_show,conDB,3,3
	'計算資料總數	
	totalput=rstObj1.recordcount
	if totalput=0 then
	else
	redim pa_delwkid(totalput)   '刪除之工作id
		'列出資料項目
		rstobj1.MoveFirst
		for j=1 to totalput
			pa_delwkid(j-1)=rstObj1.fields("tmp_id")         '刪除之工作id
			'移到下一筆記錄
			rstObj1.MoveNext
			if rstObj1.EOF=True then exit for
		next
	end if
	'關閉資料集
	rstObj1.Close
	'重設資料變數 
	set rstObj1=Nothing
	'關閉資料庫
	conDB.Close
	'重設物件變數
	set conDB=Nothing
'==================================
if totalput>0 then              '==================================
   ' 連結Access資料庫daywork.mdb
   DBpath=Server.MapPath("./database/daywork.mdb")
   strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
   '建立資料庫連結物件
   set conDB= Server.CreateObject("ADODB.Connection")
   '連結資料庫	
   conDB.Open strCon
   '開啟資料表名稱
   tb_name="work_data"
      for kj=1 to totalput          '==================================
			 	'建立資料庫存取物件	
				set rstObj1=Server.CreateObject("ADODB.Recordset")
				strSQL_show="Select * from " & tb_name & " where wk_id =" & pa_delwkid(j-1)
				rstObj1.open strSQL_show,conDB,3,3  	
					'計算資料總數	
					jt=rstObj1.recordcount
				'關閉資料集
				rstObj1.Close
				'重設資料變數 
				set rstObj1=Nothing
				if jt=1 then				      	
					'刪除資料之SQL指令字串
					strSQL_del="Delete from " & tb_name & " where wk_id =" & pa_delwkid(j-1)
					'執行SQL指令
					conDB.Execute strSQL_del
				end if
      next                        '==================================
   '關閉資料庫
   conDB.Close
   '重設物件變數
   set conDB=Nothing
end if                            '==================================


%>
<html>
<head>
<title>將暫存工作存入工作資料庫中</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--讀入螢幕顯示樣板檔 base_screen_一般.css 及列印樣板檔 base_print_一般.css  -->
	<link rel="stylesheet" type="text/css" 
		media="screen" href="./css/base_screen.css" title="style_screen">
<!--設定樣板格式-->
<style type="text/css">
	<!--

	-->
</style>
</head>
<body>
<%
'自動關閉網頁
Response.Write "<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> "
%>
</body>
</html>