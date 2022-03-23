<% @codepage=950%>

<%
pyymmdd=year(date())& right("0"& month(date()),2) & right("0"& day(date()),2)
pfilename="calendar_google_"&pyymmdd&".csv"
Response.AddHeader "Content-Disposition","attachment;filename="&pfilename
Response.ContentType = "application/vnd.ms-csv"
%>
<%
'讀取查詢條件資料 
querystr=" where "
querystra=""
querystrb="查詢條件："

'主旨p_wk_item
p_wk_item=request("p_wk_item")
if p_wk_item="" or p_wk_item="不限" then
	p_wk_item="不限"
else
	querystra=querystra & "wk_item like '%"& p_wk_item &"%' and "
	querystrb=querystrb & "[主旨="& trim(p_wk_item) &"]"
	querystrc=querystrc & "p_wk_item"& trim(p_wk_item) &"&"
end if

'執行人員p_wk_exe
p_wk_exe=trim(request("p_wk_exe"))
'p_wk_exe="惟亭"	
if p_wk_exe="" or p_mtclass="不限" then
	p_wk_exe="不限"
else
	querystra=querystra & "(wk_exe like '%"& p_wk_exe &"%' or wk_exe like '全體人員' ) and "
	querystrb=querystrb & "[執行人員="&trim(p_wk_exe)&"]"
	querystrc=querystrc & "p_wk_exe="&trim(p_wk_exe)&"&"
end if


'執行日期p_doing_date1a
p_doing_date1a=trim(request("p_doing_date1a"))	
'p_doing_date1a="2016/3/1"
if p_doing_date1a="" or p_doing_date1a="不限" then
	p_doing_date1a="不限"
else
	querystra=querystra & "(doing_date1 >= #"& p_doing_date1a &"# ) and "
	querystrb=querystrb & "[執行日期="&trim(p_doing_date1a)&"]"
	querystrc=querystrc & "p_doing_date1a="&trim(p_doing_date1a)&"&"
end if

'執行日期p_doing_date1b
p_doing_date1b=trim(request("p_doing_date1b"))	
'p_doing_date1b="2016/4/1"
if p_doing_date1b="" or p_doing_date1b="不限" then
	p_doing_date1b="不限"
else
	querystra=querystra & "(doing_date1 <= #"& p_doing_date1b &"# ) and "
	querystrb=querystrb & "[執行日期="&trim(p_doing_date1b)&"]"
	querystrc=querystrc & "p_doing_date1b="&trim(p_doing_date1b)&"&"
end if

	querystr=querystr & querystra
	len_a=len(querystr)
	if len_a=7 then querystr=" "
      if trim(querystr)="where" then querystr=" "
	if right(querystr,4)="and " then querystr=left(querystr,len_a-4)
	len_c=len(querystrc)
	if right(querystrc,1)="&" then querystrc=left(querystrc,len_c-1)
	
%>
<%
'輸出資料直接轉為utf-8 65001
Response.Charset="utf-8"
Session.Codepage=65001
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
'strSQL_show="Select * from " & tb_name & " where wk_class like '%"&wk_class&"%' and wk_doer like '%"& wk_man &"%' order by doing_date1 desc"
strSQL_show="Select * from " & tb_name & querystr &" order by doing_date1 asc"
'Response.Write querystr & vbCrLf
rstObj1.open strSQL_show,conDB,1,1
totalput=rstObj1.recordcount
%>

<%
if totalput=0 then
	str_00="Subject"'活動名稱 (必要)。
	str_01="Start Date"'活動的第一天 (必要)。
	str_02="Start Time"'活動開始時間。
	str_03="End Date"'活動的最後一天。
	str_04="End Time"'活動結束時間。
	str_05="All Day Event"'這個活動是否為全天活動。如果是全天活動，請輸入 True；否則請輸入 False。
	str_06="Description"'活動說明或附註。
	str_07="Location"'活動地點。
	str_08="Private"'這個活動是否為私人活動。如果是私人活動，請輸入 True；否則請輸入 False。
	Response.Write str_00 & "," & str_01 & "," & str_02 & "," & str_03 & "," & str_04 & "," & str_05 & "," & str_06 & "," & str_07 & "," & str_08 & vbCrLf
else
	str_00="Subject"'活動名稱 (必要)。
	str_01="Start Date"'活動的第一天 (必要)。
	str_02="Start Time"'活動開始時間。
	str_03="End Date"'活動的最後一天。
	str_04="End Time"'活動結束時間。
	str_05="All Day Event"'這個活動是否為全天活動。如果是全天活動，請輸入 True；否則請輸入 False。
	str_06="Description"'活動說明或附註。
	str_07="Location"'活動地點。
	str_08="Private"'這個活動是否為私人活動。如果是私人活動，請輸入 True；否則請輸入 False。
	Response.Write str_00 & "," & str_01 & "," & str_02 & "," & str_03 & "," & str_04 & "," & str_05 & "," & str_06 & "," & str_07 & "," & str_08 & vbCrLf

	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
	'讀取資料
		doing_date1=rstObj1.fields("doing_date1")'派工日期
		wk_item=replace(trim(rstObj1.fields("wk_item")),",","，")'主旨
		wk_item=replace(wk_item,";",":")'主旨
		wk_content=left(rstObj1.fields("wk_content"),200)'工作內容註記
		wk_content=replace(wk_content,",","，")'工作內容註記
		wk_content=replace(wk_content,chr(13),"。")'工作內容註記
		wk_content=replace(wk_content,chr(10),"")'工作內容註記
		str1_02a=left(wk_item,5)
		if not(isnumeric(left(str1_02a,2))) then
			str1_02a="08:00"
		end if
		str1_04a=Mid(wk_item,7,5)
		if not(isnumeric(left(str1_04a,2))) then
			str1_04a=str1_02a
		end if
	str1_00=wk_item	'活動名稱 (必要)。
	str1_01=doing_date1		'活動的第一天 (必要)。
	str1_02=str1_02a	'活動開始時間。
	str1_03=doing_date1		'活動的最後一天。
	str1_04=str1_04a		'活動結束時間。
	str1_05="False"				'這個活動是否為全天活動。如果是全天活動，請輸入 True；否則請輸入 False。
	str1_06=wk_content	'活動說明或附註。
	str1_07="taipei"'活動地點。
	str1_08="False"'這個活動是否為私人活動。如果是私人活動，請輸入 True；否則請輸入 False。
	Response.Write str1_00 & "," & str1_01 & "," & str1_02 & "," & str1_03 & "," & str1_04 & "," & str1_05 & "," & str1_06 & "," & str1_07 & "," & str1_08 & vbCrLf

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
%>
<%
'輸出資料直接轉為utf-8 65001
Response.Charset="big-5"
Session.Codepage=950
%>