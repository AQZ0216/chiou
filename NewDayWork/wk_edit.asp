<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<!-- Include file = "./include/array_worker.inc" -->
<!-- #Include file = "./include/array_wkclass.inc" -->
<!-- #Include file = "./include/array_pjn.inc" -->
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<%
stra_gp6=""
stra_gp5=""
'人員名稱
	'stra_gp5 內業人員
'	st_dp5="財務部,法務部,資訊部,管理部"
	'stra_gp6 業務人員
'	st_dp6="企劃部,業務部"
'	st_dp7="基金會"
for ki=1 to worker_no
	if staff_gp_a(ki-1)="內業" then
		stra_gp5= stra_gp5 & "," & worker_a(ki-1)
	elseif left(staff_gp_a(ki-1),1)="業" then
		stra_gp6= stra_gp6 & "," & worker_a(ki-1)
'	elseif staff_gp_a(ki-1)="社企" then
'		stra_gp7= stra_gp7 & "," & worker_a(ki-1)
	end if
	stra_gp0= stra_gp0 & "," & worker_a(ki-1)
next
'response.write "worker_no="&worker_no&"。<br>"
'response.write "stra_gp6="&stra_gp6&"。<br>"
'response.write "stra_gp5="&stra_gp5&"。<br>"
'response.write "stra_gp7="&stra_gp7&"。<br>"
'response.end
stra_gp0=right(stra_gp0,len(stra_gp0)-1) '全體
stra_gp6=right(stra_gp6,len(stra_gp6)-1) '業務人員
stra_gp5=right(stra_gp5,len(stra_gp5)-1) '內業人員
'stra_gp7=right(stra_gp7,len(stra_gp7)-1) '社企
stra_gp1="郭董,國賢,國哲,少鑫,維尼,美慧,惠娟,惟亭,寶元"   '郭董行程專用
%>

<%
'工作等級陣列
'dim wk_class_a
'wk_class_a=array("","A","B","C","D","Z")
'wk_class_no=ubound(wk_class_a)+1
%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
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
wk_class1=rstObj1.fields("wk_class")
wk_group1=rstObj1.fields("wk_group")
wk_exe1=rstObj1.fields("wk_exe")
wk_att=rstObj1.fields("wk_att")
wk_pjid=rstObj1.fields("pj_id")          '專案名稱id
wk_pjn=rstObj1.fields("pj_02")          '專案名稱
pwk_password=rstObj1.fields("wk_password")   '加密文字
wk_headline=rstObj1.fields("headline")'跑馬燈

if wk_group1="專案工作" and doing_date1 < date() then doing_date1=date()

%>
<%
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
body{font-family:'微軟正黑體';background-color:'#F0FFF0'}
input{font-family:'微軟正黑體';}
textarea{font-family:'微軟正黑體';}
SELECT{font-family:'微軟正黑體';font-size:12pt;}
td{font-family:'微軟正黑體';}
--></style>
</HEAD>
<BODY>
<center>
<form name="form1" action="wk_edit_ok.asp" method="post">
<input type=hidden name="wk_id1" value="<%=wk_id%>">
<!-- #Include file = "./include/toolbar_edit.inc" -->
<% 
	if trim(wk_order)=trim(worker) then
%>
<!-- #Include file = "./include/wk_edit_form_all.inc" -->
<% 
	else
%>
<!-- #Include file = "./include/wk_edit_form.inc" -->
<% 
	end if
%>
</form>
<script language=vbscript>
<%
for i=1 TO worker_no
%>
sub worker_s<%=i%>_click()
	if document.form1.wk_doer.value="" then
		document.form1.wk_doer.value=Trim(document.form1.worker_s<%=i%>.value)
		document.form1.wk_undoer.value=Trim(document.form1.worker_s<%=i%>.value)
	else
	    document.form1.wk_doer.value=document.form1.wk_doer.value &","& Trim(document.form1.worker_s<%=i%>.value)
	    document.form1.wk_undoer.value=document.form1.wk_undoer.value &","& Trim(document.form1.worker_s<%=i%>.value)
	end if
end sub
<%
next
%>
sub all_sel_click()
	document.form1.wk_doer.value=""
	<%
	for i=1 TO worker_no
	%>	
		worker_s<%=i%>_click
	<%
	next
	%>	
end sub
sub all_unsel_click()
	document.form1.wk_doer.value=""
end sub
'sub exeadd()
'  if document.form1.exemen_w.value="clear" then
'   document.form1.wk_exe.value=""
'  else
'	if document.form1.wk_exe.value="" then
'		document.form1.wk_exe.value=document.form1.exemen_w.value
'	else
'         if instr(1,document.form1.wk_exe.value,document.form1.exemen_w.value,1)>0 then
'            document.form1.wk_exe.value=replace(document.form1.wk_exe.value,document.form1.exemen_w.value,"")
'            document.form1.wk_exe.value=replace(document.form1.wk_exe.value,",,",",")
'            if left(document.form1.wk_exe.value,1)="," then document.form1.wk_exe.value=right(document.form1.wk_exe.value,len(document.form1.wk_exe.value)-1)
'            if right(document.form1.wk_exe.value,1)="," then document.form1.wk_exe.value=left(document.form1.wk_exe.value,len(document.form1.wk_exe.value)-1)
'         else
'		document.form1.wk_exe.value=document.form1.wk_exe.value & "," & document.form1.exemen_w.value
'         end if
'	end if
'  end if
'	document.form1.exemen_w.value=""
'end sub
</script>
<center>
</body>
</html>
