<%@ Language=VBScript CODEPAGE=950 %>
<%

'每頁顯示筆數為100筆 data_no
data_no=100
'目前頁碼 page_no
if request("page_no")="" then
	page_no=1 
else
	page_no=request("page_no")
end if

%>
<!-- 開啟資料庫 -->
<%
' 連結Access資料庫./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="project_data"

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " order by pj_01 desc"
rstObj1.open strSQL_show,conDB,3,3
'計算資料總數	
totalput=rstObj1.recordcount	
%>	 

<html>
<head>
<title>專案列表</title>
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

<!-- 標題 -->
<%
if totalput=0 then
%>
<font style="font-family:新細明體;font-size:5mm;font-weight:bold;color:#006400;">
無專案！！
</font> 
<%
else
	data_ck=totalput mod data_no
	if data_ck=0 then
		page_total=int(totalput/data_no)
	else
		page_total=int(totalput/data_no)+1
	end if
%>
<%
if page_total=1 then
		page_no_b=1
		page_no_g=1
else
	if page_no=1 then
		page_no_b=1
		page_no_g=page_no+1
	else
		if cint(page_no)=cint(page_total) then
			page_no_b=page_no-1
			page_no_g=page_total
		else
			page_no_b=page_no-1
			page_no_g=page_no+1
		end if
	end if
end if

'計算起始筆數
datafirst=(page_no-1)*data_no+1
if cint(page_no)=cint(page_total) then
	datalast=totalput
else
	datalast=datafirst+data_no-1
end if
'計算本頁筆數 
no_local=datalast-datafirst+1
'設定session backURL
strbackURLa="pj_list.asp?page_no="
strbackURL=strbackURLa&page_no
Session("strbackURL")=strbackURL

%>
<font style="font-family:新細明體;font-size:3.0mm;font-weight:normal;">
<table border=0 style="width:750px;">
<tr style="height:35px;">
<td style="width:20%;text-align:center;background-color:#e0e0e0;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#e0e0e0';">
<a href="pj_add.asp" target="_self" style="text-decoration:none;color:blue;">【新增專案】</a>
</td> 
<td style="width:60%;text-align:center;">
<font style="font-family:新細明體;font-size:4mm;font-weight:bold;color:#006400;">
所有《專案編號》列表
</font>
<%
for j=1 to page_total
%>
<a href="<%=strbackURLa%><%=j%>" ><%=j%>&nbsp;</a>
<%
next
%>
<td style="width:20%;text-align:center;background-color:#e0e0e0;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#e0e0e0';">
<a href="pj_list_print.asp?page_no=<%=page_no%>" target="_self" style="text-decoration:none;color:blue;">【列印本頁】</a>
</td> 
</table>
</font>

<div style="text-align:left;width:775px;height:55px;overflow:off;">
<!-- 資料列表標題開始 -->
<font style="font-family:新細明體;font-size:3.5mm;font-weight:normal;">
<table border=0 style="font-size:3.5mm;text-align:left;width:750px;">
<tr>
<a href="<%=strbackURLa%><%=page_no_b%>">
	<td style="width:8%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		<a href="<%=strbackURLa%><%=page_no_b%>">前一頁<br>(第<%=page_no_b%>頁)</a>
	</td>
</a>
	<td style="font-size:4mm;width:8%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		第<font color=red><%=page_no%></font>頁
	</td>
<a href="<%=strbackURLa%><%=page_no_g%>">
	<td style="width:8%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		<a href="<%=strbackURLa%><%=page_no_g%>">下一頁<br>(第<%=page_no_g%>頁)</a>
	</td>
</a>
	<td style="text-align:center;">
		<font style="font-family:新細明體;font-size:3.5mm;font-weight:normal;">
		目前頁碼是第<font color=red>&nbsp;<%=page_no%>&nbsp;</font>頁，共有<font color=red>&nbsp;<%=page_total%>&nbsp;</font>頁(每頁<%=data_no%>筆)，共有<font color=red><%=totalput%></font>筆資料<br> 
		本頁資料為第<font color=red>&nbsp;<%=datafirst%>&nbsp;至&nbsp;<%=datalast%>&nbsp;</font>筆，共<font color=red><%=no_local%></font>筆資料
		</font>
<a href="<%=strbackURLa%><%=page_no_b%>">
	<td style="width:10%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		<a href="<%=strbackURLa%><%=page_no_b%>">前一頁<br>(第<%=page_no_b%>頁)</a>
	</td>
</a>
	<td style="font-size:4mm;width:8%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		第<font color=red><%=page_no%></font>頁
	</td>
<a href="<%=strbackURLa%><%=page_no_g%>">
	<td style="width:10%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		<a href="<%=strbackURLa%><%=page_no_g%>">下一頁<br>(第<%=page_no_g%>頁)</a>
	</td>
</a>
</table>
</font>
<font style="font-family:新細明體;font-size:4mm;font-weight:normal;">
<!-- 資料列表畫面內容 -->
<input type='hidden' name="firstdata" value="<%=datafirst%>">
<input type='hidden' name="lastdata" value="<%=datalast%>">
<input type='hidden' name="local_no" value="<%=no_local%>">
<!-- 資料列表內容開始 -->
<font style="font-family:新細明體;font-size:3.5mm;font-weight:normal;">
	<table border=0 style="text-align:left;width:750px;">
	<col style="font-size:3.5mm;width:60px;text-align:center;">
	<col style="font-size:3.5mm;width:80px;text-align:center;">
	<col style="font-size:3.5mm;width:120px;text-align:center;">
	<col style="font-size:3.5mm;;text-align:center;">
	<tr>
		<td>功能選擇</td> 
		<td>專案編號
		<td>專案名稱 
		<td>專案說明 
	</tr>
	</table>
</font>
</div>
<div style="text-align:left;width:775px;height:255px;overflow:auto;">
<font style="font-family:新細明體;font-size:3.5mm;font-weight:normal;">
	<table border=0 style="text-align:left;width:750px;">
	<col style="font-size:3.5mm;width:60px;text-align:center;">
	<col style="font-size:3.5mm;width:80px;text-align:center;">
	<col style="font-size:3.5mm;width:120px;text-align:center;">
	<col style="font-size:3.5mm;;text-align:left;">

<%

'移至第一筆資料 
rstobj1.MoveFirst
'移至起始筆數 
rstobj1.move datafirst-1

%>
	<%	
	'列出資料項目
	'rstobj1.MoveFirst
	for j=datafirst to datalast
	'設定空白資料之反映
p_id=rstObj1.fields("pj_id")	'專案id
p_01=rstObj1.fields("pj_01")	'專案編號
p_02=rstObj1.fields("pj_02")	'專案名稱
p_03=rstObj1.fields("pj_03")	'專案說明

oddchk = j mod 2
if oddchk=1 then
	BKC="#ddffee"
else
	BKC="#ffffee"
end if
	%>
	<tr style="background-color:<%=BKC%>;" onmouseover="javascript:this.style.background='#FFeedd';" onmouseout="javascript:this.style.background='<%=BKC%>';">
	<td valign=middle >
	<a href="./pj_del.asp?p_id=<%=p_id%>"> <img src="./img/del1.gif" alt="刪除專案" width="15" height="15" style='cursor:hand;border:0;'></a>
	<a href="./pj_edit.asp?p_id=<%=p_id%>"> <img src="./img/edit1.gif" alt="編輯專案" width="15" height="15" style='cursor:hand;border:0;'></a>
	</td>
	<td><%=p_01%></td>
	<td><a href="./pj_show.asp?p_id=<%=p_id%>" ><%=p_02%></a></td>
	<td>&nbsp;&nbsp;<%=p_03%></td>
</tr>	
	
	<%
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
<!-- 資料列表結束-->	
</table>
</font>
</div>

</body>
</html> 

