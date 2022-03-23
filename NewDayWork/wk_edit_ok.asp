<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
	p_wk_content=trim(request("wk_content"))
	p_wk_item=trim(request("wk_item"))
	p_doing_date1=request("doing_date1")
	p_wk_class=request("wk_class")      '工作分類

	p1_wk_exe=request("wk_exe")
	p_wk_att=request("wk_att")
	'p_wk_checker=request("wk_checker")     '完成人員
	p_wk_undoer=request("wk_undoer")     '未完成人員
	p_wk_doer=request("wk_doer")       '知會人員
	p_redo=request("redo")  '重新通知修改
	if p_redo="是" then p_wk_item=p_wk_item&" [★"&date()&"修改★]"                   
if  instr(1,p_wk_doer,worker,1)=0 then p_wk_doer=p_wk_doer&","&worker	

p_wk_pjn=request("wk_pjn")          '專案名稱

'if p_wk_pjn="0" or isnull(p_wk_pjn) then
'      p_pj_id=0
'      p_pj_02=null
'elseif  p_wk_pjn="" then
if trim(p_wk_pjn)="，" then
      p_pj_id=null
      p_pj_02=null
else
      a_wk_pjn=split(p_wk_pjn,"，",-1,1)
      p_pj_id=a_wk_pjn(0)
      p_pj_02=a_wk_pjn(1)
end if

p_wk_password=request("str_pwd")      '加密文字
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
	'----------2003/03/15修正 
	'修改資料之SQL指令字串 全部資料
	'strSQL_edit="Update "&tb_name&" set wk_content='"&request("wk_content")&"'"
	'strSQL_edit=strSQL_edit & ",doing_date1=#"& request("doing_date1") &"#"
	'strSQL_edit=strSQL_edit & ",wk_item='"& request("wk_item") &"'"
	'strSQL_edit=strSQL_edit & " where wk_id =" & wk_id
	'執行SQL指令
	'conDB.Execute strSQL_edit
	'---------------------------------------------------------
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id="&wk_id
rstObj1.open strSQL_show,conDB,1,3
rstobj1.MoveFirst
'讀取資料
po_wk_doer=rstObj1.fields("wk_doer")     '舊知會人員
po_wk_undoer=rstObj1.fields("wk_undoer")               '舊未完成人員
po_checker=rstObj1.fields("wk_checker")                               '舊已完成人員
po_group=rstObj1.fields("wk_group")    '工作群組
'修改資料
rstObj1.fields("wk_content")= trim(p_wk_content)            '內容
rstObj1.fields("doing_date1")= p_doing_date1                   '執行日期
rstObj1.fields("wk_item")= p_wk_item                                '主旨

if  po_group="一般工作" then
   rstObj1.fields("wk_class")= p_wk_class                               '分類
else
   rstObj1.fields("pj_02")= p_pj_02                               '分類
   rstObj1.fields("pj_id")= p_pj_id                               '分類
end if

rstObj1.fields("wk_exe")= p1_wk_exe         '執行人員
rstObj1.fields("wk_att")= p_wk_att         '出席人員
rstObj1.fields("wk_doer")= p_wk_doer     '知會人員

'判斷新知會人員--------------------------------------------------------------------------------
pn_wk_doer=p_wk_doer        '新知會人員
pa_wk_doer=split(po_wk_doer,",",-1,1)
pa_wk_doer_no=ubound(pa_wk_doer)+1
for pai=1 to pa_wk_doer_no
   pn_wk_doer=replace(pn_wk_doer,pa_wk_doer(pai-1),"")
   pn_wk_doer=replace(pn_wk_doer,",,",",")
next
if left(pn_wk_doer,1)="," then pn_wk_doer=right(pn_wk_doer,len(pn_wk_doer)-1)
if right(pn_wk_doer,1)="," then pn_wk_doer=left(pn_wk_doer,len(pn_wk_doer)-1)

'將新知會人員加入未完成人員中-------------------------------------
if pn_wk_doer="" then
   pn_wk_undoer=po_wk_undoer
else
   if po_wk_undoer="" then
      pn_wk_undoer=pn_wk_doer
   else
      pn_wk_undoer=po_wk_undoer&","&pn_wk_doer
   end if
end if
rstObj1.fields("wk_undoer")=pn_wk_undoer
'將新知會人員加入未完成人員中-------------------------------------

if p_redo="是" then
   rstObj1.fields("wk_undoer")=p_wk_doer
   rstObj1.fields("wk_checker")=""
else
	rstObj1.fields("wk_undoer")=p_wk_undoer
end if

rstObj1.fields("wk_password")=p_wk_password

rstObj1.UpdateBatch
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數
set conDB=Nothing 

strURL1="wk_show.asp?wk_id="&wk_id
response.redirect(strURL1)
%>

</body>
</html>
