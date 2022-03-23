<%@ Language=VBScript CODEPAGE=950 %>
<%
'20120517更新 ========
   '讀取人員姓名
   worker = Session("worker")
   wk_id=Request("wk_id")

%>
<html>
<head>
<title>資料修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'建立資料庫存取物件  
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,3 
'讀取資料

wk_content=rstObj1.fields("wk_content")
wk_doer=rstObj1.fields("wk_doer")                    '工作人員
wk_checker=rstObj1.fields("wk_checker")           '已檢查人員(所有工作人員)
wk_undoer=rstObj1.fields("wk_undoer")               '未完成工作人員
wk_finisher=trim(rstObj1.fields("wk_finisher"))      '完成人員(填寫完成之人員)

'人名在完成名單中移除
if isnull(wk_finisher) or wk_finisher="" then
   '人名在完成名單中移除
   wk_checker=replace(wk_checker,worker,"")
   wk_checker=replace(wk_checker,",,",",")
   if left(wk_checker,1)="," then
      wk_checker=replace(wk_checker,",","",1,1)
   end if
else
   '人名在完成名單中移除
   wk_checker=replace(wk_checker,worker,"")
   wk_checker=replace(wk_checker,",,",",")
   if left(wk_checker,1)="," then
      wk_checker=replace(wk_checker,",","",1,1)
   end if
end if

'將人名在未完成工作者之名單加入
if wk_undoer="" or isnull(wk_undoer) then
   wk_undoer=worker
else
   wk_undoer=worker & "," & wk_undoer 
end if
'在工作內容中增加完成日期及人名
wk_content=wk_content & chr(13) & worker & "於" & date() &"取消完成工作"

rstObj1.fields("wk_content")=wk_content
rstObj1.fields("done_date1")=done_date1
rstObj1.fields("wk_checker")=wk_checker
rstObj1.fields("wk_finisher")=wk_finisher
rstObj1.fields("wk_undoer")=wk_undoer
rstObj1.UpdateBatch


'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
%>

<%
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing

'strbackURL=session("strbackURL")
strbackURL="wk_show.asp?wk_id="&wk_id
response.redirect(strbackURL)

%>
</body>
</html>
