<%@ Language=VBScript CODEPAGE=950 %>
<%
'函數檢查工作是否存在
function exist_wkid(pwk_id)
      ' 連結Access資料庫daywork.mdb
      DBpath_fe=Server.MapPath("./database/daywork.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '建立資料庫連結物件
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '連結資料庫	
      conDB_fe.Open strCon_fe
      '開啟資料表名稱
      tb_name_fe="work_data"
      '建立資料庫存取物件	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where wk_id = "& pwk_id &" order by wk_id desc"
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
      exist_wkid=totalput_fe
end function
%>
<HTML>
<HEAD>
<title>所有上傳之附加檔案列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<%

'附加檔案列表
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/attach_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="file_data"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where del_ok = false order by wk_id desc, fl_date desc"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
else

%>
<table border=1 cellspacing=0 cellpadding=0 width=750 >
<col width=40 style="text-align:center;">
<col width=280 style="padding-left:5px;text-align:left;">
<col width=210 style="padding-left:5px;text-align:left;">
<col width=90 style="text-align:center;">
<col width=90 style="text-align:center;">
<tr>
<td colspan=5 style="font-size:15pt;color:blue;">所有附加檔案列表</td>
</tr>
<tr>
<td >序號</td>
<td align=center >檔案說明</td>
<td align=center >檔案名稱 [上傳者]</td>
<td >建檔日期</td>
<td >功能</td>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for fi=1 to totalput
	'讀取資料
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")
		pfl_item=rstObj1.fields("fl_item")
		pfl_inputer=rstObj1.fields("fl_inputer")
		pfl_history= rstObj1.fields("fl_history")
		pfl_date=rstObj1.fields("fl_date")
		str_none=pwk_id&"_"
		str_pfl_name=right(pfl_name,len(pfl_name)-len(pwk_id)-1)
%>
<tr>
<td ><%=fi%></td>
<td ><a href="./1_ulf_item_edit.asp?fl_id=<%=pfl_id%>" target="_self" title="修改檔案說明。" ><img src="./img/change.png" style="vertical-align:middle;height:16px;cursor:hand;border:0;" ></a>
<%=pfl_item%></td>
<td ><a href="./file_att/<%=pfl_name%>" target="_blank" title="<%=pfl_history%>"><%=str_pfl_name%></a> [<%=pfl_inputer%>]</td>
<td ><%=pfl_date%></td>
<td >
<% if exist_wkid(pwk_id)=1 then %>
<input type="button" name="shfile" value="工"  onclick="file_sh('<%=pwk_id%>')" title="顯示原工作項目 [ wk_id=<%=pwk_id%> ] 內容。">
<input type="button" name="addfile" value="新"  onclick="file_add('<%=pwk_id%>')" title="工作項目 [ wk_id=<%=pwk_id%> ] 新增檔案。">
<% end if %>
<input type="button" name="delfile" value="刪"  onclick="file_del('<%=pfl_id%>')" title="將檔案刪除。">
<!-- <a href="1_ulf_form.asp?wk_id=<%=pwk_id%>" title="新增檔案或更新檔案。">新</a>
<a href="1_ulf_del.asp?fl_id=<%=pfl_id%>" title="刪除檔案。">刪</a> -->
</td>
</tr>
<%
	'移到下一筆記錄
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>

</table>
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
</form>
</center>
<script language=vbscript>
sub file_sh(ppwk_id)
	ok=msgbox("是否確定要顯示工作項目？"&chr(13)&"wk_show.asp?wk_id="&ppwk_id,1,"新增警告")
	if ok=1 then 
		'location.href="wk_show.asp?wk_id="&ppwk_id
		window.open("wk_show.asp?wk_id="&ppwk_id)
	end if
end sub
sub file_add(ppwk_id)
	ok=msgbox("是否確定要新增檔案？"&chr(13)&"1_ulf_form.asp?wk_id="&ppwk_id,1,"新增警告")
	if ok=1 then 
		location.href="1_ulf_form.asp?wk_id="&ppwk_id
	end if
end sub
sub file_del(ppfl_id)
	ok=msgbox("是否確定要刪除檔案？"&chr(13)&"1_ulf_del.asp?fl_id="&ppfl_id,1,"刪除警告")
	if ok=1 then 
		location.href="1_ulf_del.asp?fl_id="&ppfl_id
	end if
end sub

</script>
</body>
</html>