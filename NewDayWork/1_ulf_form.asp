<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
%>
<%
'BASP21.DLL將檔案上載程式
'請先執行安裝RegSvr32 Basp21.dll
'可將表單中的text也抓出塞進陣列，直拉response.write 變數，就可以print出來了
'-------------------------------------------------------------------
'上傳附件檔案畫面
wk_id=request("wk_id") '讀取工作項目之wk_id

if wk_id="" or isnull(wk_id ) then wk_id=0

%>

<HTML>
<HEAD>
<Title>上傳檔案畫面</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<form id="form1" name="form1" method="post" action="1_ulf_form_ok.asp" enctype="multipart/form-data">
<input type="hidden" name="text" value="<%=wk_id%>" >
<table width=760 border=0 cellspacing=0 cellpadding=0 bgcolor="#FFFFBB">
<col width=60>
<col width=240>
<col width=60>
<col width=200>
<col width=80>
<tr>
<td colspan=5 align=center>
<b>上傳工作項目附件檔案</b>
<a href="wk_show.asp?wk_id=<%=wk_id%>" title="工作wk_id=<%=wk_id%>">回工作內容</a>
</td>
</tr>
<tr>
   <td align=right>檔案說明</td>
   <td><input type="text" name="item" value="" style="width:100%" maxlength="40"></td>
   <td align=right>檔案名稱</td>
   <td><input type="file" name="image" style="width:100%"></td>
   <td><input type="submit" name="button1" value="上傳檔案" style="width:100%"></td>
</tr>
<tr>
<td colspan=5 align=left style="padding-left:5px;">
注意：同一工作如果上傳相同<font color=blue>檔案名稱</font>時，將會取代原檔案及說明。
</td>
</tr>
</table>

<%
'附加檔案列表
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/attach_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="file_data"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id &" and del_ok = false"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
   response.write "無任何上傳附件。"
else
%>
<table border=1 cellspacing=0 cellpadding=0 width=750 bgcolor="#CCEEFF">
<col width=40 style="text-align:center;">
<col width=340 style="padding-left:5px;text-align:left;">
<col width=260 style="padding-left:5px;text-align:left;">
<col width=100 style="text-align:center;">
<tr>
<td colspan=4>現有上傳附件列表</td>
</tr>
<tr>
<td >序號</td>
<td align=center >檔案說明</td>
<td align=center >檔案名稱 [上傳者]</td>
<td >建檔日期</td>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for fi=1 to totalput
	'讀取資料
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")
		pfl_item=rstObj1.fields("fl_item")
		pfl_inputer=rstObj1.fields("fl_inputer")
		pfl_history= rstObj1.fields("fl_history")
		pfl_date=rstObj1.fields("fl_date")
		str_none=pwk_id&"_"
		str_pfl_name=right(pfl_name,len(pfl_name)-len(pwk_id)-1)
%>
<tr>
<td ><%=fi%></td>
<td ><%=pfl_item%></td>
<td ><a href="./file_att/<%=pfl_name%>" target="_blank" title="<%=pfl_history%>"><%=str_pfl_name%></a> [<%=pfl_inputer%>]</td>
<td ><%=pfl_date%></td>
</tr>
<%
	'移到下一筆記錄
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>

</table>
<%
end if
'關閉資料集
rstObj1.Close
'重設資料變數
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數
set conDB=Nothing
%>
<!--
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
<%
function showspace(ztxt)
   if ztxt="" or isnull(ztxt) then
      pztxt="&nbsp;"
   else
      pztxt=ztxt
   end if
   showspace=pztxt
end function
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<tr>
	<td align="center" colspan=2 rowspan=2><font size=4 color="red"><b>顯示單一工作表</b></font></td>
	<td align="right">工作群組：</td>
	<td><%=showspace(wk_group)%>
	</td>
	<td align="right">專案名稱：</td>
	<td><%=showspace(wk_pjn)%>
	</td>
</tr>

<tr>
	<td align="right">工作編號：</td>
	<td><%=showspace(wk_id)%>
	</td>
	<td align="right">工作分類：</td>
	<td><%=showspace(wk_class)%>
	</td>
</tr>

<tr>
	<td align="right">公告者：</td>
	<td><%=showspace(wk_order)%>
	</td>
	<td align="right">公告日期：</td>
	<td><%=showspace(undo_date1)%>
	</td>
	<td align="right">執行日期：</td>
	<td><%=showspace(doing_date1)%>
	</td>
</tr>
<tr>
	<td align="right">
	知會人員：
	</td>
	<td colspan=5><%=showspace(wk_doer)%>
	</td>
</tr>
<tr>
	<td align="right">
	完成人員：
	</td>
	<td colspan=5><%=showspace(wk_checker)%>
	</td>
</tr>
<tr>
	<td align="right">
	未完成人員：
	</td>
	<td colspan=5><%=showspace(wk_undoer)%>
	</td>
</tr>
<tr>
	<td align="right">
	主旨：
	</td>
	<td colspan=5><%=showspace(wk_item)%>
	</td>
</tr>
</table>
-->
<hr>
</form>
</center>
</BODY>
</HTML>