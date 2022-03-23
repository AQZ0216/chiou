<% @codepage=950%>
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
	pwd_headline=Request("pwd_headline")   '密碼
%>

<%
if pwd_headline="3939" then
      '將工作列為重大訊息
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
      strSQL_show="Select * from " & tb_name & " where wk_id="&wk_id
      rstObj1.open strSQL_show,conDB,1,3
      rstObj1.fields("headline")=5
      rstObj1.UpdateBatch
      '關閉資料集
      rstObj1.Close
      '重設資料變數
      set rstObj1=Nothing
      '關閉資料庫
      conDB.Close
      '重設物件變數
      set conDB=Nothing
    strURL1=session("hback_URL")
      'strURL1="wk_lst_doing.asp"
      response.redirect(strURL1)

else
    if isnull(pwd_headine) or pwd_headline="" then
         pwd_msg=""
    else
         pwd_msg="密碼錯誤！！"
    end if
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
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'讀取資料
undo_date1=rstObj1.fields("undo_date1")
doing_date1=rstObj1.fields("doing_date1")
done_date1=rstObj1.fields("done_date1")
wk_item=rstObj1.fields("wk_item")
wk_content=rstObj1.fields("wk_content")
wk_order=rstObj1.fields("wk_order")
wk_doer=rstObj1.fields("wk_doer")
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_class=rstObj1.fields("wk_class")
wk_group=rstObj1.fields("wk_group")
wk_exe=rstObj1.fields("wk_exe")
wk_pjn=rstObj1.fields("pj_02")   '專案名稱
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing
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
<center>
<form name="form1" action="0_wk_headline_off.asp" method="post">
<input type="hidden" name="wk_id" value="<%=wk_id%>">
<input type="hidden" name="worker" value="<%=worker%>">
<font style="font-size:16pt;" color="red">要將此重大訊息取消，請輸入密碼！！</font><br>
<font style="font-size:12pt;" color="blue">將【主旨】在首頁跑馬燈中取消！！</font><br>
密碼：<input type='password' name='pwd_headline' value='' style="width:100px;" > <br>
		<input type=submit name="editok" value="確定" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100px;">
		<input type=button name="goback1" value="回上一頁" onclick="history.back()" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100px;" >
<hr>
<%
if pwd_msg="密碼錯誤！！" then
      response.write "<b>" & pwd_msg & "</b><hr>"
end if
%>

<table border=1 cellspacing=0 cellpadding=0>
<col width=120><col width=120><col width=120><col width=120><col width=120><col width=120>
<tr>
	<td align="center" colspan=2 rowspan=2><font size=4 color="red"><b>顯示單一工作表</b></font></td>
	<td align="right">工作群組：</td>
	<td>&nbsp;<%=wk_group%></td>
	<td align="right">專案名稱：</td>
	<td>&nbsp;<%=wk_pjn%></td>
</tr>
<tr>
	<td align="right">工作編號：</td>
	<td>&nbsp;<%=wk_id%></td>
	<td align="right">工作分類：</td>
	<td>&nbsp;<%=wk_class%></td>
</tr>
<tr>
	<td align="right">公告者：</td>
	<td>&nbsp;<%=wk_order%></td>
	<td align="right">公告日期：</td>
	<td>&nbsp;<%=undo_date1%></td>
	<td align="right">執行日期：</td>
	<td>&nbsp;<%=doing_date1%></td>
</tr>
<tr>
	<td align="right">
	知會人員：
	</td>
	<td colspan=5>&nbsp;<%=wk_doer%></td>
</tr>
<tr>
	<td align="right">
	完成人員：
	</td>
	<td colspan=5>&nbsp;<%=wk_checker%></td>
</tr>
<tr>
	<td align="right">
	未完成人員：
	</td>
	<td colspan=5>&nbsp;<%=wk_undoer%></td>
</tr>
<tr>
	<td align="right">
	主旨：
	</td>
	<td colspan=5>&nbsp;<%=wk_item%></td>
</tr>
<tr>
	<td align="right" valign="top">
	執行內容：
	</td>
	<td colspan=5>
	<%
	 wk_content=replace(wk_content,chr(13),"<br>")
	 response.write wk_content
	%>
	</td>
</tr>
</table>
</form>
<center>
</body>
</html>
<% end if %>