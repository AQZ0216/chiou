<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_id=Request("wk_id")
	wk_chk=Request("wk_chk")
	strbackURL=session("strbackURL")
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
wk_class=rstObj1.fields("wk_class")
wk_group=rstObj1.fields("wk_group")
wk_exe=rstObj1.fields("wk_exe")           '執行人員
wk_att=rstObj1.fields("wk_att")           '出席人員
wk_pjn=rstObj1.fields("pj_02")   '專案名稱
pwk_password=rstObj1.fields("wk_password")   '加密文字
wk_headline=rstObj1.fields("headline")'跑馬燈

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
' <%
' '判斷是否是IE或手機
' dim u,b
' set u=Request.ServerVariables("HTTP_USER_AGENT")
' 'response.write u
' 'response.write "<hr>"
' 'response.end
' '
' ck_MSIE=instr(1,u,"MSIE",1)
' ck_IE=instr(1,u,"IE",1)

' ck_Chrome=instr(1,u,"Chrome",1)
' ck_Firefox=instr(1,u,"Firefox",1)

' ck_Safari=instr(1,u,"Safari",1)
' ck_Firefox=instr(1,u,"Firefox",1)

' if ck_MSIE+ck_IE>0 then
'    'IE瀏覽器
'    ck_mobile=0
' elseif ck_Chrome >0 then
'    'Chrome瀏覽器
'    ck_mobile=1
' elseif ck_Firefox>0 then
'    'Firefox瀏覽器
'    ck_mobile=1
' elseif ck_Safari>0 then
'    'Safari瀏覽器
'    ck_mobile=1
' else
'    ck_mobile=1
' end if

' if ck_mobile=1 then
'       nexturl="3_mobilejs_wk_show.asp?wk_id="& wk_id&"&wk_chk="&wk_chk
'       response.redirect(nexturl)
' else
' end if

' 'set b=new RegExp
' 'b.Pattern="firefox|chrome|safari|mobile"
' 'b.Pattern="safari|mobile"
' 'b.IgnoreCase=true
' 'b.Global=true
' 'Set matchesb = b.Execute(u)
' 'if b.test(u) then               '非IE瀏覽器
' '      response.redirect("http://detectmobilebrowser.com/mobile")
' '      response.write "b="& matchesb(0).value &"<hr>"
' '      response.write "b.test(u)="&b.test(u)&"<hr>"
' '      response.write "瀏覽器："& matchesb(0).value & "<hr>"
'       '非IE
'       'nexturl="3_mobilejs_wk_show.asp?wk_id="& wk_id&"&wk_chk="&wk_chk
'       'response.redirect(nexturl)
' 'else
' '      response.write "b.test(u)="&b.test(u)&"<hr>"
' '      response.write "瀏覽器："&"IE<hr>"
' 'end if
' 'response.end
' %>
<%
'是否加密文件
if wk_chk="ok" or pwk_password="" or isnull(pwk_password) then
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
<form name="form1" action="" method="post">
<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<input type="hidden" name="worker1" value="<%=worker%>">
<input type="hidden" name="wk_ordera" value="<%=wk_order%>">
<%if wk_group="一般工作" then%>
<!-- Include file = "./include/toolbar_show.inc" -->
<script language=vbscript>
'sub edit_click()
'	wk_id=document.form1.wk_id1.value
'	location.href="./wk_edit.asp?wk_id="&wk_id
'end sub
sub delete_click()
	worker=document.form1.worker1.value
	wk_order=document.form1.wk_ordera.value
	if worker=wk_order then
		ok=msgbox("是否確定要刪除資料？",1,"刪除警告")
		if ok=1 then
			wk_id=document.form1.wk_id1.value
			location.href="./wk_del_ok.asp?wk_id="&wk_id
		end if
	else
		ok=msgbox("你不是派工者，無法刪除此項工作！！",0,"錯誤警告")
	end if
end sub
sub done_click()
	ok=msgbox("是否確定要完成工作？",1,"確認警告")
	if ok=1 then
		wk_id=document.form1.wk_id1.value
		location.href="./wk_done_ok.asp?wk_id="&wk_id
	end if
end sub
sub readd_click()
	ok=msgbox("是否確定要重新公告工作？",1,"確認警告")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_readd.asp?wk_id="&wk_id
	end if
end sub
sub gpchange_click()
	ok=msgbox("是否確定要轉為專案工作？",1,"確認警告")
	if ok=1 then
		wk_id=document.form1.wk_id1.value
		location.href="./wk_gpchg_special.asp?wk_id="&wk_id
	end if
end sub
</script>

<script type="text/javascript">
function edit_click()
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./wk_edit.asp?wk_id="+x1;
   // alert(str_url);
   location.href = str_url ;
}

</script>

<center>

<table border=0 cellspacing=0 cellpadding=0>
<col span=8 style="width:90px;text-align:center;">
<tr width=720>	
	<td>
		<input type=button name="bkpg" value="回上一頁" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="parent.location.href='javascript:history.back()'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="edit" value="編修工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="edit_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="delete" value="刪除工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="delete_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
<% if wk_class="Z" then %>
		<input type=button name="done" value="完成工作" disabled style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
<% else %>
		<input type=button name="done" value="完成工作" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="done_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
<% end if %>
	</td>
	<td>
		<input type=button name="readd" value="重新公告" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="readd_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="wkprint" value="列印內容" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="window.open('./wkprint_si.asp?wk_id=<%=wk_id%>')"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="gpchange" value="轉為專案" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="gpchange_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="wkattfile" value="上傳附件" style="cursor:hand;background-color:'#77FFEE';color:blue;width:100%;" onclick="location.href='./1_ulf_form.asp?wk_id=<%=wk_id%>'"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#77FFEE';">
	</td>
</tr>
</table>
</center>
<%else%>
<!-- #Include file = "./include/toolbar_pj_show.inc" -->
<%end if%>
<!-- #Include file = "./include/wk_show_form.inc" -->

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
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id &" and del_ok = false order by fl_date desc"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
else
%>
<table border=1 cellspacing=0 cellpadding=0 width=750 bgcolor="#CCEEFF">
<col width=40 style="text-align:center;">
<col width=280 style="padding-left:5px;text-align:left;">
<col width=210 style="padding-left:5px;text-align:left;">
<col width=90 style="text-align:center;">
<col width=90 style="text-align:center;">
<tr>
<td colspan=5>附件列表</td>
</tr>
<tr>
<td >序號</td>
<td align=center >檔案說明</td>
<td align=center >檔案名稱  [上傳者]</td>
<td >建檔日期</td>
<td >功能</td>
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
<td >
<a href="./1_ulf_item_edit.asp?fl_id=<%=pfl_id%>" target="_self" title="修改檔案說明。" ><img src="./img/change.png" style="vertical-align:middle;height:16px;cursor:hand;border:0;" ></a>
<%=pfl_item%>
</td>
<td ><a href="./file_att/<%=pfl_name%>" target="_blank" title="<%=pfl_history%>"><%=str_pfl_name%></a>  [<%=pfl_inputer%>]</td>
<td ><%=pfl_date%></td>
<td >
<input type="button" name="addfile" value="新"  onclick="file_add('<%=pwk_id%>')" title="工作項目 [ wk_id=<%=pwk_id%> ] 新增檔案。">
<input type="button" name="delfile" value="刪"  onclick="file_del('<%=pfl_id%>')" title="將檔案刪除。">
<!-- <a href="1_ulf_form.asp?wk_id=<%=pwk_id%>" title="新增檔案或更新檔案。">新</a>
<a href="1_ulf_del.asp?fl_id=<%=pfl_id%>" title="刪除檔案。">刪</a> -->
</td>
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
</form>
<hr>
<!-- 瀏覽器資料：<%=u%> -->
</center>
<script language=vbscript>
sub file_add(ppwk_id)
	ok=msgbox("是否確定要新增資料？"&chr(13)&"1_ulf_form.asp?wk_id="&ppwk_id,1,"新增警告")
	if ok=1 then 
		location.href="1_ulf_form.asp?wk_id="&ppwk_id
	end if
end sub
sub file_del(ppfl_id)
	ok=msgbox("是否確定要刪除資料？"&chr(13)&"1_ulf_del.asp?fl_id="&ppfl_id,1,"刪除警告")
	if ok=1 then 
		location.href="1_ulf_del.asp?fl_id="&ppfl_id
	end if
end sub

</script>
</body>
</html>
<%
else     '加密文件輸入加密文字
%>
<html>
<head>
<title>密碼檢查</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'新細明體';background-color :'#FFFEEE'}
input{
	font-family:'新細明體';
	font-size:12pt;
	}
select{font-family:'新細明體';font-size:10pt;cursor:hand;}
.itxt{
	font-family:'新細明體';
	font-size:12pt;
	width:100%;
	height:100%;
	}
input.imenu { 
	/*font-size:15px;				/*字體大小*/
	/*font-weight:bold;
	cursor:hand;				/*游標形式*/ 
	background-color:'<%=botton_color%>'; 		
	margin:0 0 0 0;		/*邊緣上下左右*/
	width:100px;
	/*height:100%;*/
	color:#000000;
	letter-spacing:2px;
	cursor:hand;
     }
td{
	margin:0 0 0 0;		/*邊緣上下左右*/
	border-color:'#808080'; /*表格外框顏色*/ 
	border-style:solid;		/*表格外框線型*/
	border-width:1px;		/*表格外框厚度*/  
	vertical-align:middle;	/*字體垂直對齊方式*/
	/*font-size:15px;*/ 
	}
table{	
	margin:0 0 0 0;		/*邊緣上下左右*/
	border-collapse:collapse; 	/*邊框形式重合*/
	}
input.itext { 
	font-size:3.5mm;				/*字體大小*/
	/*cursor:hand;				/*游標形式*/ 
	width:100%;
	height:5mm;
	background-color:'#ffeedd'; 		/*外框顏色*/
	margin:0 0 0 0;		/*邊緣上下左右*/
	color:black;
	text-align:right;
     }

--></style>
</head>
<body>
<center>
<form name="form_login" method=post action="">
<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<input type="hidden" name="wk_pwd" value="<%=pwk_password%>">
<table border=0 cellspacing=0 cellpadding=0>
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<tr>
	<td align="center" colspan=2 ><font size=4 color="red"><b>顯示單一工作表</b></font></td>
	<td align="right">工作群組：</td>
	<td><!-- <%=showspace(wk_group)%> -->
	<input type='text' name='wk_group' value='<%=wk_group%>' style="width:100%;" readonly>
	</td>
	<td align="right">專案名稱：</td>
	<td><!-- <%=showspace(wk_pjn)%> -->
 	<input type='text' name='wk_pjn' value='<%=wk_pjn%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">公告者：</td>
	<td><!-- <%=showspace(wk_order)%>-->
	<input type='text' name='wk_order1' value='<%=wk_order%>' style="width:100%;" readonly>
	</td>
	<td align="right">公告日期：</td>
	<td><!-- <%=showspace(undo_date1)%> -->
	<input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly>
	</td>
	<td align="right">執行日期：</td>
	<td><!--<%=showspace(doing_date1)%>-->
 	<input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	執行人員：
	</td>
	<td colspan=5><!--<%=showspace(wk_checker)%> -->
 	<input type='text' name='wk_exe' value='<%=wk_exe%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right" style="background-color:#FFBFFF;">
	出席人員：
	</td>
	<td colspan=5>
 	<input type='text' name='wk_att' value='<%=wk_att%>' style="width:100%;" readonly  onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right">
	知會人員：
	</td>
	<td colspan=5><!--<%=showspace(wk_doer)%> -->
 	<input type='text' name='wk_doer' value='<%=wk_doer%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	完成人員：
	</td>
	<td colspan=5><!--<%=showspace(wk_checker)%> -->
 	<input type='text' name='wk_checker' value='<%=wk_checker%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	未完成人員：
	</td>
	<td colspan=5><!--<%=showspace(wk_undoer)%> -->
 	<input type='text' name='wk_undoer' value='<%=wk_undoer%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	主旨：
	</td>
	<td colspan=5><!--<%=showspace(wk_item)%>-->
 	<input type='text' name='wk_item' value='<%=wk_item%>' style="width:100%;" readonly>
	</td>
</tr>
</table>
<hr color=red>
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">【請輸入加密文字以檢視全部內容】</font>
<table border=0 cellspacing=0 cellpadding=2 style="width:300px" >
<col style="width:100px;font-size:4mm;" align=center>
<col style="width:200px;font-size:4mm;" align=center>
<tr>
<td>加密文字：</td>
<td><input type="text" style="text-align:left;" name="wkr_pwd" value="" maxlength="10" ></td>
</tr>
<tr>
<td colspan=2>
	<input type="button" name="submit01" value="檢視全部內容" onclick="check_password()">
	<input type="button" name="reset01" value="回上頁" onclick="back_url()" >
</td>
</tr>
</table>
<hr color=red>
</body>
</html>
<script language='Vbscript'>
<!--
sub check_password()
   chk_str=document.form_login.wk_pwd.value
   ipt_str=document.form_login.wkr_pwd.value
   pwk_id=document.form_login.wk_id1.value
   if chk_str=ipt_str then
      str_url="./wk_show.asp?wk_id="&pwk_id&"&wk_chk=ok"
      'MyVar = MsgBox ("wk_id="&pwk_id&"。chk_str="&chk_str, 16, "錯誤訊息")
      location.href=str_url
   else
      str_url="<%=strbackURL%>"
      MyVar = MsgBox ("加密文字錯誤！！"&chr(13)&"回到上一頁！！", 16, "錯誤訊息")
      location.href=str_url
   end if

end sub
sub back_url()
   back_str="<%=strbackURL%>"
   location.href=back_str
end sub
-->
</script>
<%
end if
%>