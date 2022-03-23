<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker.inc" -->
<!-- #Include file = "./include/array_wkclass.inc" -->
<!-- #Include file = "./include/array_wkgroup.inc" -->
<!-- #Include file = "./include/workinput.inc" -->
<!-- #Include file = "./misc_data/array_place.inc" -->	
<!-- #Include file = "./misc_data/array_thing.inc" -->	
<!-- #Include file = "./misc_data/array_writer.inc" -->	
<%
	'讀取人員姓名
	worker = Session("worker")
	datecode=request("datecode")
	wk_order=worker
	undo_date1=date()
'工作等級陣列 
'dim wk_class_a
'wk_class_a=array("","A","B","C","D")
'wk_class_no=ubound(wk_class_a)+1
%>
<HTML>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body {  scrollbar-3dlight-color:#ffffff;
        scrollbar-arrow-color:#CCCCCC;
        scrollbar-base-color:#666633;
        scrollbar-darkshadow-color:#e6e6cc;
        scrollbar-face-color:#666666;
        scrollbar-highlight-color:#ffffff;
        scrollbar-shadow-color:#e6e6cc;
        scrollbar-track-color:#ffffff;
        margin:2mm 0mm 0mm 0mm;		/*邊緣上下左右*/
		font-family:'標楷體';		/*字形*/
		font-size:4mm; 			/*字體大小*/
		background-color:'#F0FFF0';
     }
input.imenu { 
	font-size:3.5mm;				/*字體大小*/
	cursor:hand;				/*游標形式*/ 
	background-color:'#d3d3d3'; 		/*外框顏色*/ 
	margin:0 0 0 0;		/*邊緣上下左右*/
     }
input.imenu1 { 
	font-size:3.5mm;	/*字體大小*/
	font-weight:bold;				
	cursor:hand;				/*游標形式*/ 
	background-color:'#eeeeff'; 		/*外框顏色*/ 
	margin:0 0 0 0;		/*邊緣上下左右*/
	width:80px;
	height:100%;
     }
     
TD.SOME{
		font-family: '標楷體';
		font-size: 3.3mm;
		line-height: 18px;
		color:blue;
		font-weight:bold;
		}
TD.myd{
		font-family: '標楷體';
		font-size: 3.3mm;
		line-height: 18px;
		background-color:#f0ffff;
		}     
    
-->
</style>

</HEAD>
<BODY>
<center>

<form name="form1" action="wk_add_ok.asp" method="post" >
<input type="hidden" name="worker1" value="<%=worker%>" >
<table border=1 cellspacing=0 cellpadding=0>
<col width=100>
<col width=130>
<col width=100>
<col width=130>
<col width=100>
<col width=130>
<tr>
	<td colspan=4 align=center><font size=4 color="red"><b><%=worker%>工作公告單</b></font></td>
	<td align="right"><font color="red">工作群組：</font></td>
	<td>
	<select name="wk_group" style="width:100%;height:100%;font-size:10pt;background-color:'#ffffee';">
<%
		response.write "<option value='"&wk_group_a(0)&"' selected>"&wk_group_a(0)
	for i=2 to wk_group_no
		response.write "<option value='"&wk_group_a(i-1)&"'>"&wk_group_a(i-1)
	next
%>
	</select>
	</td>
<tr>
	<td align="right">公告者：</td>
	<td><input type='text' name='wk_order' value='<%=wk_order%>' style="width:100%;" readonly></td>
	<td align="right">公告日期：</td>
	<td><input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly></td>
	<td align="right"><font color="red">工作分類：</font></td>
	<td>
	<select name="wk_class" style="width:100%;height:100%;font-size:10pt;background-color:'#ffffee';" onchange="item_chk()">
<%
	for i=1 to wk_class_no
		response.write "<option value='"&wk_class_a(i-1)&"'>"&wk_class_a(i-1)
	next
%>
	</select>
	</td>
</tr>
<tr>
	<td colspan=6 align="center">
<table border="0" cellspacing="0" cellpadding="0">
<col width=120><col width=120><col width=120><col width=120><col width=120>
<tr><td align="center" valign="middle">
	<!-- 人員選項 -->
	<font size=3>人員選項：<br>
		<SELECT name="men_w" onchange="menadd()">
		<option selected>請選擇人員</option>
	<%
		for i=1 to writer_no
			response.write "<option value='"&writer_a(i-1)&"'>"&writer_a(i-1)
		next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- 日期選項 -->
	<font size=3>日期選項：<br>
		<img align=top src="img/cal3.gif" onmousedown="Cal('doing_date1')" width="20" height="20" style='cursor:hand;'>
	</td>
	<td align="center" valign="middle">
	<!-- 時間選項 -->
	<font size=3>時間選項：<br>
		<SELECT name="time_w" onchange="timeadd()">
		<option selected>請選擇時間</option>
	<%
	for i=1 to 19
		Response.Write("<OPTION value=" & wk_time(i-1) & ">" & wk_time(i-1)&"</OPTION>")
	next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- 地點選項 -->
	<font size=3>地點選項：<br>
		<SELECT name="place_w" onchange="placeadd()">
		<option selected>請選擇地點</option>
	<%
		for i=1 to place_no
			response.write "<option value='"&place_a(i-1)&"'>"&place_a(i-1)
		next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- 事件選項 -->
	<font size=3>事件選項：<br>
		<SELECT name="thing_w" onchange="thingadd()">
		<option selected>請選擇事件</option>
	<%
		for i=1 to thing_no
			response.write "<option value='"&thing_a(i-1)&"'>"&thing_a(i-1)
		next
	%>
		</font></SELECT>
	</td>

</tr>
</table>
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">主旨：</font>
	</td>
	<td colspan=3>
	<input type='text' name='wk_item' value='' style="width:100%;" onchange="item_chk()">
	</td>
	<td align="right"><font color="red">執行日期：</font></td>
	<td><input type='text' name="doing_date1" value="<%=datecode%>" style="width:70%;">
		<img align=top src="img/cal3.gif" onmousedown="Cal('doing_date1')" width="20" height="20" style='cursor:hand;'>
	</td>

</tr>

<tr>
	<td align="right">
	<font color="red">執行內容：</font>
	</td>
	<td colspan=5>
	<TEXTAREA name="wk_content" rows="5" style="width:100%;" ><%=wk_content%></TEXTAREA>
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">知會人員：</font>
	</td>
	<td colspan=5>
	<input type='text' name='all_worker' value='<%=worker%>' size='85' onchange="item_chk()">
	</td>
</tr>
<tr>
	<td colspan=6 align="center">
<%	
	for i=1 to worker_no
%>
	<input type="button"  name="worker_s<%=i%>" value="<%=worker_a(i-1)%>" onclick="worker_s<%=i%>_click()">
<%
	next
%>	
	<input type="button" name="all_sel" value="全部人員" onclick="all_sel_click()">
	<input type="button" name="all_unsel" value="清除人員" onclick="all_unsel_click()">
	</td>
</tr>
<tr>
	<td colspan=6 align="center">
	<input type="button" name="press" value="確定公告" onclick="press_chk()">
	<input type="reset" name="cancel" value="清除資料" >
	</td>
<tr>
</table>
</form>

<!--月曆產生的位置-->
<Span ID=ShowCal style="position:absolute;z-index:1;"></Span>

<Script Language=VBScript>
<!--
Sub Cal(TObject)

'產生月曆，以今天的日期為基準
 Call GetCal(Year(Now()),Month(Now()),TObject)

'調整<Span>的位置
 ShowCal.style.left=window.event.clientX-140
 ShowCal.style.top=window.event.clientY+10
End Sub 

Sub GetCal(SYear,SMonth,TObject)
'月曆抬頭部分
Str=Str &"<table ALIGN='CENTER' BORDER='0' CELLSPACING='0' CELLPADDING='2' BGCOLOR='#f0ffff' BORDERCOLOR='Gray'>"
Str=Str &"<tr><td>"
Str=Str &"        <table WIDTH='140' BORDER='0' CELLPADDING='1' CELLSPACING='0' BGCOLOR='#FFFFFF'>"
Str=Str &"                <tr HEIGHT='18' BGCOLOR='Silver'>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/prev3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='上一月' style='cursor:hand' OnClick='PreMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/Next3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='下一月' style='cursor:hand' OnClick='NextMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='100' COLSPAN='4' ALIGN='CENTER' VALIGN='MIDDLE' CLASS='SOME'><Span ID=SelYear>" & SYear & "</span>年<Span ID=SelMon>" & SMonth & "</Span>月</td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='RIGHT' VALIGN='MIDDLE'><img SRC='img/cdia3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='關閉視窗' style='cursor:hand;' OnClick='Closedate()' ></td>"
Str=Str &"                </tr>"
Str=Str &"          <tr HEIGHT='15' BGCOLOR='Aliceblue'>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>日</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>一</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>二</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>三</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>四</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>五</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>六</td>"
Str=Str &"          </tr>"
Str=Str &"      <tr>"

'該月第一天之星期
SDate=DateValue(SYear & "/" & SMonth & "/1")
SWeek=WeekDay(SDate)

'該月最後一天的日期
EDay=Day(DateSerial(SYear,SMonth+1,0))
EDate=DateValue(SYear & "/" & SMonth & "/" & EDay)

'該月最後一天之星期
EWeek=WeekDay(EDate)

'上月最後一天的日期
PreEDay=Day(DateSerial(SYear,SMonth,0))

'產生上個月的部分
Dim i
For i=1 to SWeek-1
 Str=Str & "<td CLASS='myd' width=20 align=right><font color=gray>" & PreEDay-SWeek+i+1 & "</font></td>"
Next

If SWeek=1 then
 Str=Str & "</tr>"
End if

'產生本月部分
SW=SWeek
i=1
For i=1 to EDay
 '調整六日字型的顏色
 Select Case SW
 Case 1
  FColor="Red"
 Case 7
  FColor="Green"
 Case Else
  FColor="Black"
 End Select
  
  GDate=SYear & "/" & SMonth & "/" & i
  Str=Str & "<td CLASS='myd' width=20 align=right style='cursor:hand;' onMouseOver=""this.style.backgroundColor='#FF99FF'"" onMouseOut=""this.style.backgroundColor='#f0ffff'"" Onclick=""SendDate('" & GDate & "','" & TObject & "')""><Font Color=" & FColor & ">" & i & "</Font></td>"
 
'產生下個月部分
SW=SW+1
 IF SW>7 then
  Str=Str & "</tr><tr>"
  SW=1
 End if
Next

J=1
For i=SW to 7
 Str=Str & "<td CLASS='myd' width=20 align=right><Font Color=Gray>" & j & "</Font></td>"
 J=j+1
Next


Str=Str & "      </tr>"
Str=Str & "</Table>"
'將資料引入<Span>
ShowCal.InnerHTML=Str

End Sub

'前移一個月
Sub PreMon(TObject)
 SYear=Int(SelYear.OuterTEXT)
 SMon=int(SelMon.outerTEXT)-1

 '判斷是否往前調一年
 IF SMon<1 then
  SMon=12
  SYear=SYear-1
 End if
 Call GetCal(SYear,SMon,TObject)
End Sub

'後移一個月
Sub NextMon(TObject)
 SYear=Int(SelYear.OuterTEXT)
 SMon=int(SelMon.outerTEXT)+1

 '判斷是否往前往一年
 IF SMon>12 then
  SMon=1
  SYear=SYear+1
 End if
 Call GetCal(SYear,SMon,TObject)
End Sub

'將資料送入欄位內
Sub SendDate(GDate,TObject)
 document.all.namedItem(TObject).Value=GDate
 	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.all.namedItem(TObject).Value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.all.namedItem(TObject).Value
	end if
 ShowCal.InnerHTML=""
End Sub

'關閉日期畫面
Sub CloseDate()
 ShowCal.InnerHTML=""
End Sub
-->
</script>


<script language=vbscript>
<%
for i=1 TO worker_no
%>
sub worker_s<%=i%>_click()
	if document.form1.all_worker.value="" then
		document.form1.all_worker.value=Trim(document.form1.worker_s<%=i%>.value)
	else
		document.form1.all_worker.value=document.form1.all_worker.value &","& Trim(document.form1.worker_s<%=i%>.value)
	end if
end sub
<%
next
%>
sub all_sel_click()
	document.form1.all_worker.value=""
	<%
	for i=1 TO worker_no
	%>	
		worker_s<%=i%>_click
	<%
	next
	%>	
end sub
sub all_unsel_click()
	document.form1.all_worker.value=document.form1.worker1.value
end sub
</script>
<script language=vbscript>
sub menadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.men_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.men_w.value
	end if
end sub
sub dateadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.date_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.date_w.value
	end if
	'document.form1.doing_date1.value=document.form1.date_w.value
end sub
sub timeadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.time_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.time_w.value
	end if
end sub
sub placeadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.place_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.place_w.value
	end if
end sub
sub thingadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.thing_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.thing_w.value
	end if
end sub
sub item_chk()
	if document.form1.wk_item.value="" or document.form1.all_worker.value="" or document.form1.wk_class.value="" then
		ok=msgbox("請輸入工作分類、主旨及知會人員！！",0,"錯誤警告")
	else
	end if
end sub
sub press_chk()
	if document.form1.wk_item.value="" or document.form1.all_worker.value="" or document.form1.wk_class.value="" then
		ok=msgbox("請輸入工作分類、主旨及知會人員！！",0,"錯誤警告")
	else
		form1.submit
	end if
end sub
</script>

</center>
</body>
</html>
