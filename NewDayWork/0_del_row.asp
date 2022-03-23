<!-- 開啟資料庫 -->
<%
'設定讀取資料編號
last_row=Request("last_row")    'p_id
'=======================================================================
' 連結Access資料庫./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="linkdata"
'刪除資料之SQL指令字串
strSQL_del="Delete from " & tb_name & " where lk_row =" & last_row
'執行SQL指令
conDB.Execute strSQL_del
'關閉資料庫
conDB.Close
'重設物件變數 
set conDB=Nothing
'=======================================================================

nexturl="firstpage_elink.asp"
response.redirect(nexturl)
%>	
<html>
<head>
<title>確定刪除資料</title>
<style type="text/css"><!--
body{font-family:'新細明體';background-color :'#FFFEEE'}
input{font-family:'新細明體';font-size:12pt;cursor:hand;}
select{font-family:'新細明體';font-size:10pt;cursor:hand;}
--></style>
</head>
<body>
<center></center>
</body>
</html>