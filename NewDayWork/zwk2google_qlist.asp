<%@ Language=VBScript CODEPAGE=950 %>

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
if p_wk_exe="" or p_wk_exe="不限" then
	p_wk_exe="不限"
else
	querystra=querystra & "(wk_exe like '%"& p_wk_exe &"%') and "
'	querystra=querystra & "(wk_exe like '%"& p_wk_exe &"%' or wk_exe like '全體人員' ) and "
	querystrb=querystrb & "[執行人員="&trim(p_wk_exe)&"]"
	querystrc=querystrc & "p_wk_exe="&trim(p_wk_exe)&"&"
end if

'知會人員p_wk_doer
p_wk_doer=trim(request("p_wk_doer"))
'p_wk_doer="惟亭"	
if p_wk_doer="" or p_wk_doer="不限" then
	p_wk_doer="不限"
else
	querystra=querystra & "(wk_doer like '%"& p_wk_doer &"%') and "
	querystrb=querystrb & "[知會人員="&trim(p_wk_doer)&"]"
	querystrc=querystrc & "p_wk_doer="&trim(p_wk_doer)&"&"
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
	
if trim(querystrc)="" or isnull(trim(querystrc)) then querystrc="p_wk_item=不限"
	qstrURL="zwk2google_qlist.asp?"&querystrc
'設定session backURL
strbackURLcsv="zwk2google_qlist_csv.asp?"&querystrc
%>
<HTML>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<%
'response.write "strbackURLcsv="&strbackURLcsv
'response.write "<hr>"
'response.write "<a href='"& strbackURLcsv &"' target='_blank'>google csv</a>"
'response.write "<hr>"
%>
<form name="form1" method="post" action="zwk2google_qlist_csva.asp" >
	<input type="button" name="sentb" class="cbutton" value="確定匯出" onclick="Verify_chk()" >
	<input type="reset" name="reset" class="cbutton" value="清除資料"  >
	<input type=button name=giveup class="cbutton" value="回上一頁" onclick="history.back()"  >
<hr>
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
<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'strSQL_show="Select * from " & tb_name & " where wk_class like '%"&wk_class&"%' and wk_doer like '%"& wk_man &"%' order by doing_date1 desc"
strSQL_show="Select * from " & tb_name & querystr &" order by doing_date1 asc"
'	response.write strSQL_show &"<br>"
rstObj1.open strSQL_show,conDB,1,1
totalput=rstObj1.recordcount
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
'	Response.Write str_00 & "," & str_01 & "," & str_02 & "," & str_03 & "," & str_04 & "," & str_05 & "," & str_06 & "," & str_07 & "," & str_08 & "<br>"
	Response.Write "無資料可匯出。"
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
'	Response.Write str_00 & "," & str_01 & "," & str_02 & "," & str_03 & "," & str_04 & "," & str_05 & "," & str_06 & "," & str_07 & "," & str_08 & "<br>"
%>
<table border=1 style="width:1000px;">
	<col style="width:30px;text-align:center;">
	<col style="width:160px;text-align:center;">
	<col style="width:80px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:80px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:;text-align:center;">
	<col style="width:100px;text-align:center;">
	<col style="width:60px;text-align:center;">
<tr>
	<td align=center colspan=10>共有<%=totalput%>筆資料可匯出。
	</td>
</tr>	
<tr>
	<td align=center width=30>序號
		<input type="checkbox" name="psel_wkid" value="" onclick="sel_check()" title="全選或全不選"></td>
	<td align=center >Subject</td>
	<td align=center >Start Date</td>
	<td align=center >Start Time</td>
	<td align=center >End Date</td>
	<td align=center >End Time</td>
	<td align=center >All Day Event</td>
	<td align=center >Description</td>
	<td align=center >Location</td>
	<td align=center >Private</td>
</tr>	
</table>
<div style="text-align:left;width:1020px;height:305px;overflow:auto;">
<table border=1 width=1000>
	<col style="width:30px;text-align:center;">
	<col style="width:160px;text-align:left;">
	<col style="width:80px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:80px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:;text-align:left;">
	<col style="width:100px;text-align:center;">
	<col style="width:60px;text-align:center;">
<%
	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
	'讀取資料
		wkid=rstObj1.fields("wk_id")
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
'	Response.Write str1_00 & "," & str1_01 & "," & str1_02 & "," & str1_03 & "," & str1_04 & "," & str1_05 & "," & str1_06 & "," & str1_07 & "," & str1_08 & "<br>"
%>
<tr>
	<td><!--序號--><%=i%>
			<input type="checkbox" name="p_wkid" value="<%=wkid%>" ><%'=wkid%>
		</td>
	<td><!--Subject--><%=str1_00%></td>
	<td><!--Start Date--><%=str1_01%></td>
	<td><!--Start Time--><%=str1_02%></td>
	<td><!--End Date--><%=str1_03%></td>
	<td><!--End Time--><%=str1_04%></td>
	<td><!--All Day Event--><%=str1_05%></td>
	<td><!--Description--><%=left(str1_06,50)%></td>
	<td><!--Location--><%=str1_07%></td>
	<td><!--Private--><%=str1_08%></td>
</tr>	
<%
	'移到下一筆記錄
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>
</table>
</div>
<%
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
<hr>
	<input type="button" name="sentb" class="cbutton" value="確定匯出" onclick="Verify_chk()" >
	<input type="reset" name="reset" class="cbutton" value="清除資料"  >
	<input type=button name=giveup class="cbutton" value="回上一頁" onclick="history.back()"  >
<hr>
</form>
<script language=vbscript>
sub Verify_chk()
	set celem=document.form1.getElementsByTagName("input")
	checksel1=0
	str_err=""
 	for i=0 to celem.length-1
		if celem(i).type="checkbox" and celem(i).name="p_wkid" and celem(i).checked=true then
			checksel1=checksel1+1
		end if  
	next
	if checksel1=0 then str_err=str_err&chr(13)&"請選擇要匯出之資料！！" 
	if str_err="" then
		strmsg="確定匯出資料？"& chr(13) &" 目前已選擇"& checksel1 &"筆資料！！"
		chkq=msgbox(strmsg,64+1,"確認訊息")
		if chkq=1 then
			document.form1.submit
		else
		end if
	else 	 
		errcode=msgbox("錯誤訊息！！！"& chr(13)&str_err& chr(13) ,64+0,"錯誤訊息")
	end if
end sub
sub sel_check()
	set celema=document.form1.getElementsByTagName("input")
	if document.form1.psel_wkid.checked=true then
		check_all
	else
		uncheck_all
	end if  
end sub

sub check_all()'全選
	set celem=document.form1.getElementsByTagName("input")
 	for i=0 to celem.length-1
		if celem(i).type="checkbox" and celem(i).name="p_wkid" then
			celem(i).checked=true
		end if  
	next
end sub
sub uncheck_all()'全不選
	set celem=document.form1.getElementsByTagName("input")
 	for i=0 to celem.length-1
		if celem(i).type="checkbox" and celem(i).name="p_wkid" then
			celem(i).checked=false
		end if  
	next
end sub
</script>
</body>
</html>