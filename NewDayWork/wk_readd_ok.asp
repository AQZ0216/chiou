<%@ Language=VBScript CODEPAGE=950 %>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'判斷是否輸入工作分類 
keyword=request("wk_class")
if keyword="" then 
	response.redirect("wk_add.asp")
else
end if

%>	
<html>
<head>
<title>資料完整新增</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>

<!-- #Include file = "./include/readd_wk_ok.inc" -->

<%
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%>
<script language="Javascript">
	alert("資料新增完成！！");
	location.href="wk_lst_undo.asp";
</script>

</body>
</html>
