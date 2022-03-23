<% @codepage=950%>
<%
	'讀取人員姓名
	worker = Session("worker")
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
      strSQL_show="Select * from " & tb_name & " where headline > 5  order by doing_date1 asc"
      rstObj1.open strSQL_show,conDB,1,3
totalput=rstObj1.recordcount
if totalput=0 then
else
	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
         rstObj1.fields("headline")=5
	'移到下一筆記錄
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
end if
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
<form name="form1" action="0_wk_headline_off_all.asp" method="post">
<input type="hidden" name="worker" value="<%=worker%>">
<font style="font-size:16pt;" color="red">將所有重大訊息取消，請輸入密碼！！</font><br>
<font style="font-size:12pt;" color="blue">將首頁跑馬燈中所有重大訊息取消！！</font><br>
密碼：<input type='password' name='pwd_headline' value='' style="width:100px;" > <br>
		<input type=submit name="editok" value="確定" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100px;">
		<input type=button name="goback1" value="回上一頁" onclick="history.back()" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100px;" >
<hr>
<%
if pwd_msg="密碼錯誤！！" then
      response.write "<b>" & pwd_msg & "</b><hr>"
end if
%>

</form>
<center>
</body>
</html>
<% end if %>