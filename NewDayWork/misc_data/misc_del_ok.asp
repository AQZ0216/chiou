
<%
' 連結Access資料庫misc_data.mdb
DBpath_b=Server.MapPath("./misc_data.mdb")
strCon_b="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_b
'建立資料庫連結物件
set conDB_b= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB_b.Open strCon_b
'讀取資料表名稱
s_id=(Request("add_item"))
s_word=(Request("keyword"))
'讀取欄位名稱
Select Case s_id
Case 1
	tb_name_b="writer_table"
	item="writer"
	iword="人員"	
Case 2
	tb_name_b="place_table"
	item="place"
	iword="地點"
Case 3
	tb_name_b="thing_table"
	item="thing"
	iword="事件"
Case else
	response.redirect "../error/errormisc.htm"
End Select
if Request("keyword")="" then
	response.redirect "../error/errormisc.htm"
end if

'刪除資料之SQL指令字串
strSQL_del="Delete from "&tb_name_b&" where "&item&"='" &s_word&"'"

'執行SQL指令
conDB_b.Execute strSQL_del

'關閉資料庫 
conDB_b.Close
'重設物件變數 
set conDB_b=Nothing 
%>
<html>
<head>
<title>選項資料刪除</title>
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'##FFFFcc'}
--></style>
</head>
<body >
<center>
<p><font size=5 color="#ff0000">選項刪除完成！！</font></p>
<p><font size=5 color=red>﹝<%=iword%>﹞</font><font size=5 color="#0000ff">選項中，刪除</font><font size=5 color=red>﹝<%=Request("keyword")%>﹞</font><font size=5 color="#0000ff">項目</font></p>
<table border=1><tr>
<td width=180 align=center><a href="./misc_edit.asp">回選項編修頁</a></td>
</tr></table>
</center>
</body>
</html>