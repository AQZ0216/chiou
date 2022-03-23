<%@ Language=VBScript CODEPAGE=950 %>
<%
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
pdo_date=rstObj1.fields("doing_date1")
wk_content=rstObj1.fields("wk_content")
wk_doer=rstObj1.fields("wk_doer")
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_finisher=trim(rstObj1.fields("wk_finisher"))
'if isnull(wk_finisher) then
if isnull(wk_finisher) or wk_finisher="" then
   done_date1=cstr(date())
   wk_finisher=worker
   '將人名加入完成工作者之名單中
   wk_checker=worker

else
   if isnull(rstObj1.fields("done_date1")) then
      done_date1=cstr(date())
   else
      done_date1=cstr(rstObj1.fields("done_date1"))
   end if
   '將人名加入完成工作者之名單中
   wk_checker=wk_checker&","&worker
end if

'將人名在未完成工作者之名單去除
wk_undoer=replace(wk_undoer,worker,"")
wk_undoer=replace(wk_undoer,",,",",")
if left(wk_undoer,1)="," then
   wk_undoer=replace(wk_undoer,",","",1,1)
end if
'在工作內容中增加完成日期及人名
wk_content=wk_content & chr(13) & worker & "於" & date() &"完成工作"

'20100312更新 ======== 
rstObj1.fields("wk_content")=wk_content
rstObj1.fields("done_date1")=done_date1
rstObj1.fields("wk_checker")=wk_checker
rstObj1.fields("wk_finisher")=wk_finisher
rstObj1.fields("wk_undoer")=wk_undoer
rstObj1.UpdateBatch
'20100312更新 ======== 

'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
%>
<%
'20100312更新 ======== 
'修改資料之SQL指令字串 全部資料
'strSQL_edit="Update "&tb_name&" set wk_content='"& wk_content &"'"
'strSQL_edit=strSQL_edit & ",done_date1=#"& done_date1 &"#"
'strSQL_edit=strSQL_edit & ",wk_checker='"& wk_checker &"'"
'strSQL_edit=strSQL_edit & ",wk_finisher='"& wk_finisher &"'"
'strSQL_edit=strSQL_edit & ",wk_undoer='"& wk_undoer &"'"
'strSQL_edit=strSQL_edit & " where wk_id =" & wk_id
'執行SQL指令
'conDB.Execute strSQL_edit
'20100312更新 ======== 
%>
<%
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing

'strbackURL=session("strbackURL")
nWeeksn =DatePart("ww",date()) '讀取週數
nWeeks =DatePart("ww",pdo_date) '讀取週數
nYear = Year(pdo_date)

if nWeeksn=nWeeks then
	strbackURL="wk_week_now.asp?nWeeks="&nWeeks&"&nYear="&nYear
else 
	strbackURL=session("strbackURL")
end if
response.redirect(strbackURL)

%>
<!-- <script language="Javascript">
   alert("資料修改完成！！");
//   location.href="wk_lst_doing.asp";
   location.href="wk_Calendar_all.asp";
</script> -->

</body>
</html>
