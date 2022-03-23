<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

<html>
<head>
<title>確定刪除資料</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>
<%
'刪除資料之SQL指令字串
strSQL_del="Delete from " & tb_name & " where wk_id =" & wk_id
'執行SQL指令
conDB.Execute strSQL_del
%>
<%
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing

   strbackURL=session("strbackURL")
   if strbackURL="" or isnull(strbackURL) then strbackURL="wk_calendar_all.asp"
   response.redirect(strbackURL)
%>
<!-- <script language="Javascript">
	alert("資料刪除完成！！");
	location.href="wk_lst_doing.asp";
</script> -->
</body>
</html>
