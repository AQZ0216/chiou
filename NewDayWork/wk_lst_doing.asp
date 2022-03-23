<% @codepage=950%>
<!-- #Include file = "./include/f_week_cstr.inc" -->
<%
session("hback_URL")="./wk_lst_doing.asp" '重大訊息之回復網頁
	'讀取人員姓名
	worker = Session("worker")
	'讀取今天日期
	ckdate=date()+2
wkgroup="一般工作"

wk_sort=request("wk_sort")
if wk_sort="" then
  wk_sort=session("wk_sort")
   if wk_sort="" then wk_sort=0
end if
session("wk_sort")=wk_sort
   wk_sort_1=1-wk_sort
  ' if wk_sort=0 then
      'wk_sort_1=1
  ' elseif wk_sort=1 then
      'wk_sort_1=0
   'end if
   '設定session("strbackURL")
strbackURL="wk_lst_doing.asp"
session("strbackURL")=strbackURL
%>
<%
'查詢是否有附件
Function exist_attach(pwk_id)
      ' 連結Access資料庫daywork.mdb
      DBpath_fe=Server.MapPath("./database/attach_file.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '建立資料庫連結物件
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '連結資料庫	
      conDB_fe.Open strCon_fe
      '開啟資料表名稱
      tb_name_fe="file_data"
      '建立資料庫存取物件	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where del_ok = false and wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '關閉資料集
      rstObj1_fe.Close
      '重設資料變數
      set rstObj1_fe=Nothing
      '關閉資料庫 
      conDB_fe.Close
      '重設物件變數
      set conDB_fe=Nothing
      exist_attach=totalput_fe
End Function

%>
<!-- 開啟資料庫 -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

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

<%
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'strSQL_show="Select * from " & tb_name & " where wk_doer like '%"&worker&"%' and doing_date1 <= #"&ckdate &"# and wk_undoer like '%"&worker&"%' order by doing_date1 desc"
if  wk_sort=1 then
   strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_doer like '%"&worker&"%' and doing_date1 <= #"&ckdate &"# and doing_date1 >= #"&ckdate-30 &"# and wk_undoer like '%"&worker&"%' order by doing_date1 asc, wk_item asc"
else
   strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_doer like '%"&worker&"%' and doing_date1 <= #"&ckdate &"# and doing_date1 >= #"&ckdate-30 &"# and wk_undoer like '%"&worker&"%' order by doing_date1 desc, wk_item asc"
end if
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4><%=worker%>今日(<%=ckdate-2%>)、明日(<%=ckdate-1%>)及後天(<%=ckdate%>)皆無工作事項</font>
<%
else
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=50>
<col width=130>
<col width=420>
<col width=30>
<col width=140>
<tr >
	<td colspan=5 align=center>
	<font size=4><%=worker%>今日(<%=ckdate-2%>)、明日(<%=ckdate-1%>)及後天(<%=ckdate%>)之工作事項共:<font color=red><%=totalput%></font>筆</font>
	</td>
</tr>
<tr >
	<td align=center>序號</td>
	<td align=center><a href="wk_lst_doing.asp?wk_sort=<%=wk_sort_1%>" >執行日期</a></td>
	<td align=center>主旨</td>
	<td align=center>
	<a href="0_wk_headline_off_all.asp?wk_id=<%=wk_id%>" style="text-decoration:none;color:green;" title="清除所有重大訊息！！">★</a>
	</td>
	<td align=center>
		<a href="./pj_add.asp" target="_blank"> <img src="./img/add1.gif" alt="專案新增" width="15" height="15" style='cursor:hand;border:0;'></a>
	專案名稱
		<a href="./pj_list.asp" target="_blank"> <img src="./img/list1.gif" alt="專案列表" width="15" height="15" style='cursor:hand;border:0;'></a>
	</td>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for i=1 to totalput
	'讀取資料
		wk_gp=trim(rstObj1.fields("wk_group"))
		wk_id=rstObj1.fields("wk_id")
		undo_date1=rstObj1.fields("undo_date1")
		doing_date1=rstObj1.fields("doing_date1")
		wk_item=rstObj1.fields("wk_item")
            wk_item=replace(wk_item,"平菁雲","<font color=fuchsia  >平菁雲</font>")
		wk_order=rstObj1.fields("wk_order")
		pj_id=rstObj1.fields("pj_id")
		pj_02=rstObj1.fields("pj_02")
		p_headline=rstObj1.fields("headline")
		'檢查是否有附件 exist_attach(wk_id)
               attach_no=exist_attach(wk_id)
               if attach_no=0 then
                  str_colors="color:#000000;"
               else
                  str_colors="color:#0000FF;"
               end if
               if rstObj1.fields("wk_password")="" or isnull(rstObj1.fields("wk_password")) then
               else
                  str_colors="color:#0000FF;"
               end if
		Response.Write( "<tr>")
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center style='text-align:right;padding-right:2pt;"& str_colors &"'><font size=3>" & doing_date1 &" ("&week_cstr(doing_date1)&")</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_order &"</font></td>")
		strA="<a href=wk_show.asp?wk_id="& rstObj1.fields("wk_id")&" style='letter-spacing:1.5pt;font-size:12pt;' >"
		if wk_gp="一般工作" then
			Response.Write( "<td align=left >" & strA & wk_item &"</a></td>")
		else
			strA1="<<專案工作>>"
			Response.Write( "<td align=left style='background-color:#ffff99;'>" & strA &strA1 & wk_item &"</a></td>")
		end if
%>
      <td align=center>
<% if p_headline<= 5 or isnull(p_headline) then%>
   <a href="0_wk_headline_on.asp?wk_id=<%=wk_id%>" style="text-decoration:none;color:black;" title="列為重大訊息！！">☆</a>
<% else %>
   <a href="0_wk_headline_off.asp?wk_id=<%=wk_id%>" style="text-decoration:none;color:red;" title="取消重大訊息！！">★</a>
<% end if %>
      </td>
<%		
		if pj_id="" or isnull(pj_id) then
%>		
		<td align=left><font size=3>
	<a href="./pj_add.asp?wk_id=<%=wk_id%>" target="_blank"> <img src="./img/add1.gif" alt="新增專案名稱" width="15" height="15" style='cursor:hand;border:0;'></a>
	<a href="./pj_sel.asp?wk_id=<%=wk_id%>" target="_blank"> <img src="./img/sel1.gif" alt="選擇專案名稱" width="15" height="15" style='cursor:hand;border:0;'></a>
		</font></td>
<%
		else
%>		
		<td align=left><font size=3>
	<a href="./pj_delsel.asp?wk_id=<%=wk_id%>&p_id=<%=pj_id%>" target="_blank"> <img src="./img/del1.gif" alt="移除專案名稱" width="15" height="15" style='cursor:hand;border:0;'></a>
	<a href="./pj_show.asp?p_id=<%=pj_id%>" target="_blank"><%=pj_02%></a>
		</font></td>
<%
		end if
		Response.Write( "</tr>")

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
<center>
</body>
</html>
