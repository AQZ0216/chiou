<% @codepage=950%>

<%
pyymmdd=year(date())& right("0"& month(date()),2) & right("0"& day(date()),2)
pfilename="calendar_google_"&pyymmdd&".csv"
Response.AddHeader "Content-Disposition","attachment;filename="&pfilename
Response.ContentType = "application/vnd.ms-csv"
%>
<%
'工作id p_wkid
p_wkid=request("p_wkid")
'if p_wkid="" or isnull(p_wkid) then p_wkid=""
'response.write "p_wkid="& p_wkid &"<br>"
arr_wkid=split(p_wkid,",",-1,1)
no_wkid=ubound(arr_wkid)+1
'response.write 	p_wkid
'response.end
%>
<%
'輸出資料直接轉為utf-8 65001
Response.Charset="utf-8"
Session.Codepage=65001
%>
<%
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

for kj=1 to no_wkid
	ppwkid=arr_wkid(kj-1)
	'-----------------------------------
	'建立資料庫存取物件	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name &" where wk_id ="& ppwkid &""
	'Response.Write querystr & vbCrLf
	rstObj1.open strSQL_show,conDB,1,1
	totalput=rstObj1.recordcount
	if totalput=0 then
	else
		'列出資料項目
		rstobj1.MoveFirst
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
	end if
	'關閉資料集
	rstObj1.Close
	'重設資料變數 
	set rstObj1=Nothing
	'-----------------------------------
next

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