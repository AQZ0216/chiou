<% @codepage=950%>
<!-- 開啟工作人員陣列 -->
<!-- #Include file = "./include/array_worker_crew.inc" -->
<%
	'讀取人員姓名
	worker = Session("worker")
	wk_order=worker
'工作等級陣列 
dim wk_class_a
wk_class_a=array("未分類","A","B","C","D")
wk_class_no=ubound(wk_class_a)+1
%>

<html>
<head>
<title>查詢資料</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'微軟正黑體';background-color:'#fafad2'}
input{font-family:'微軟正黑體';}
textarea{font-family:'微軟正黑體';}
SELECT{font-family:'微軟正黑體';font-size:12pt;}
td{font-family:'微軟正黑體';}
td{font-size:4.5mm;}
input.imenu { 
	font-size:4mm;				/*字體大小*/
	cursor:hand;				/*游標形式*/ 
	width:100%;
	height:100%;
	background-color:'#ffdab9'; 		/*外框顏色*/ 
	margin:0 0 0 0;		/*邊緣上下左右*/
     }
.sel1 { 
	font-size:4mm;				/*字體大小*/
	cursor:hand;				/*游標形式*/ 
	width:100%;
	height:100%;
	background-color:'#ffffee'; 		/*外框顏色*/ 
	margin:0 0 0 0;		/*邊緣上下左右*/
     }

TD.SOME{
font-family: 新細明體;
font-size: 3.5mm;
line-height: 18px;
color:blue;
font-weight:bold;
}
TD.myd{
font-family: 新細明體;
font-size: 3.5mm;
line-height: 18px;
}

--></style>
</head>
<body >
<center>
<form name="form1" method=post action="">
	<a href="./pj_list.asp" target="_blank"><font size=4 color="blue">專案列表</font></a>
<hr color=red>
<font size=4 color="blue">工作查詢條件設定</font>
<hr>
<table border=1 cellspacing=0 cellpadding=0 style="width:600px;">
<col style="width:100px;color:#ff0000;">
<col style="width:100px;">
<col style="width:100px;color:#ff0000;">
<col style="width:100px;">
<col style="">
	<tr> 
	<td align="right">工作種類：</td>
	<td>
	<select name="wk_code_t" class=sel1 >
		<option value='完成工作'>完成工作
		<option value='執行工作'>執行工作
		<option value='預計工作'>預計工作
	</select>
	</td>
	<td align="right">工作人員：</td>
	<td align="left">
		<select name="man_tcd" class=sel1 >
		<option value="不限人員">不限人員</option>
<%	
	for i=1 to worker_no
%>
		<option value="<%=worker_a(i-1)%>"><%=worker_a(i-1)%></option>
<%
	next
%>
		</select>
	</td>
	<td>
	<input class=imenu type="button" name="qclass" style="width:100%;" value="工作種類及工作人員查詢" onclick="qcode_onclick()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
	</td>
	</tr>
</table>
<hr>
<hr>
<table border=1 cellspacing=0 cellpadding=0 style="width:600px;">
<col style="width:100px;color:#ff0000;">
<col style="width:100px;">
<col style="width:100px;color:#ff0000;">
<col style="width:100px;">
<col style="">
	<tr> 
	<td align="right">工作分類：</td>
	<td>
	<select name="wk_class_t" class=sel1 >
<%
	for i=1 to wk_class_no
		response.write "<option value='"&wk_class_a(i-1)&"'>"&wk_class_a(i-1)
	next
%>
	</select>
	</td>
	<td align="right">工作人員：</td>
	<td align="left">
		<select name="man_tc" class=sel1 >
		<option value="不限人員">不限人員</option>
<%	
	for i=1 to worker_no
%>
		<option value="<%=worker_a(i-1)%>"><%=worker_a(i-1)%></option>
<%
	next
%>
		</select>
	</td>
	<td>
	<input class=imenu type="button" name="qclass" style="width:100%;" value="工作分類及工作人員查詢" onclick="qclass_onclick()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
	</td>
	</tr>
</table>
<hr>
<table border=1 cellspacing=0 cellpadding=0 style="width:600px;">
<col style="width:100px;color:#ff0000;">
<col style="width:100px;">
<col style="width:100px;color:#ff0000;">
<col style="width:100px;">
<col style="">
<tr>
	<td align="right">工作主旨：</td>
	<td align="left" >
	<input type="text" name="qtext" style="width:100%;" >
	</td>
	<td align="right">工作人員：</td>
	<td align="left">
		<select name="man_tt" class=sel1 >
		<option value="不限人員">不限人員</option>
<%	
	for i=1 to worker_no
%>
		<option value="<%=worker_a(i-1)%>"><%=worker_a(i-1)%></option>
<%
	next
%>
		</select>
	</td>
	<td align="center">
	<input class=imenu type="button" name="qitem" style="width:100%;" value="工作主旨查詢" onclick="qitem_onclick()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
	</td>	
</tr>
</table>
<hr>
<table border=1 cellspacing=0 cellpadding=0 style="width:600px;">
<col style="width:100px;">
<col style="width:100px;">
<col style="width:100px;">
<col style="width:180px;">
<col style="">
<tr>
	<td align="center">
		<select name="qkey31" class=sel1 >
		<option value="執行日期">執行日期</option>
		<option value="派工日期">派工日期</option>
		<option value="完成日期">完成日期</option>
		</select>
	</td>
	<td align="center">
		<select name="qkey32" class=sel1 >
		<option value="等於">等於</option>
		<option value="大於等於">大於等於</option>	
		<option value="小於等於">小於等於</option>
		<option value="大於">大於</option>
		<option value="小於">小於</option>
	</td>
	<td>
	<input type="text" name="qkey33" size=12 value="<%=date()%>" style="width:70px;"> <img align='top' onmousedown="Cal('qkey33')" src='img/cal3.gif' width="16" height="16" align='top' style='cursor:hand'>
	</td>
	<td align="center"><font style="font-size:3.5mm;color:#ff0000;">(日期型式：2003/12/31)</font></td>
	<td align="center">
	<input class=imenu type="button" name="query_day" value="日期查詢" onclick="qdate_onclick()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
	</td>	
</tr>
</table>

<hr>

</form>

<!--月曆產生的位置-->
<Span ID=ShowCal style="position:absolute;z-index:1"></Span>

<Script Language=VBScript>
<!--
Sub Cal(TObject)

'產生月曆，以今天的日期為基準
 Call GetCal(Year(Now()),Month(Now()),TObject)

'調整<Span>的位置
 ShowCal.style.left=window.event.clientX
 ShowCal.style.top=window.event.clientY
End Sub 

Sub GetCal(SYear,SMonth,TObject)
'月曆抬頭部分
Str=Str &"<table ALIGN='CENTER' BORDER='1' CELLSPACING='0' CELLPADDING='2' BGCOLOR='White' BORDERCOLOR='Gray'>"
Str=Str &"<tr><td>"
Str=Str &"        <table WIDTH='140' BORDER='0' CELLPADDING='1' CELLSPACING='0' BGCOLOR='#FFFFFF'>"
Str=Str &"                <tr HEIGHT='18' BGCOLOR='Silver'>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/prev3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='上一月' style='cursor:hand' OnClick='PreMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/Next3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='下一月' style='cursor:hand' OnClick='NextMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='100' COLSPAN='4' ALIGN='CENTER' VALIGN='MIDDLE' CLASS='SOME'><Span ID=SelYear>" & SYear & "</span>年<Span ID=SelMon>" & SMonth & "</Span>月</td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='RIGHT' VALIGN='MIDDLE'><img SRC='img/cdia3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='關閉視窗' style='cursor:hand' OnClick='Closedate()'></td>"
Str=Str &"                </tr>"
Str=Str &"          <tr HEIGHT='15' BGCOLOR='Aliceblue'>"
Str=Str &"                <td Colspan=7>"
Str=Str &"                 <Table Border=0>"
Str=Str &"                  <tr>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>日</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>一</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>二</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>三</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>四</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>五</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>六</td>"
Str=Str &"         </tr>"
Str=Str &"        </Table>"
Str=Str &"          </tr>"
Str=Str &"          <tr>"
Str=Str &"           <td HEIGHT='1' ALIGN='MIDDLE' COLSPAN='7'><img SRC='images/line.gif' HEIGHT='1' WIDTH='140' BORDER='0'></td>"
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
  Str=Str & "<td CLASS='myd' width=20 align=right style='cursor:hand' onMouseOver=""this.style.backgroundColor='#FF99FF'"" onMouseOut=""this.style.backgroundColor='White'"" Onclick=""SendDate('" & GDate & "','" & TObject & "')""><Font Color=" & FColor & ">" & i & "</Font></td>"
 
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
 ShowCal.InnerHTML=""
End Sub

'關閉日期畫面
Sub CloseDate()
 ShowCal.InnerHTML=""
End Sub
-->
</Script>
<script language=vbscript>
<!--
sub qclass_onclick()
	class_q=document.form1.wk_class_t.value
	man_q=document.form1.man_tc.value
	location.href="./wk_query_okc.asp?wk_class="&class_q&"&wk_man="&man_q
end sub
sub qitem_onclick()
	q_text=document.form1.qtext.value
	man_q=document.form1.man_tt.value
	location.href="./wk_query_oki.asp?q_text="&q_text&"&wk_man="&man_q
end sub
sub qdate_onclick()
	q_key31=document.form1.qkey31.value
	q_key32=document.form1.qkey32.value
	q_key33=document.form1.qkey33.value
	location.href="./wk_query_okd.asp?q_key31="&q_key31&"&q_key32="&q_key32&"&q_key33="&q_key33
end sub
sub qcode_onclick()
	code_q=document.form1.wk_code_t.value
	man_q=document.form1.man_tcd.value
	location.href="./wk_query_okcd.asp?wk_code="&code_q&"&wk_man="&man_q
end sub
-->
</script>
</center>
</body>
</html>
