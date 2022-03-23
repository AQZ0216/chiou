
<html>
<head>
<title>選項資料新增</title>
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</head>
<body >
<!-- #Include file = "./array_place_misc.inc" -->	
<!-- #Include file = "./array_thing_misc.inc" -->	
<!-- #Include file = "./array_writer_misc.inc" -->	
<script language=vbscript>
sub keycheck
	key_ck=document.misc_add.add_item.value
	select case key_ck
	case 1
		writercheck
	case 2
		placecheck
	case 3
		thingcheck
	case else
	end select
end sub

sub placecheck
	dim place_arr(<%=place_no%>)
	dim place_no
	place_no=<%=place_no%>

<%	for i=1 to place_no 
%>
		place_arr(<%=i-1%>)="<%=place_a(i-1)%>"
<%	next
%>
	for i=1 to place_no
		place_ck=Ucase(Trim(document.misc_add.keyword.value))
		if place_ck=place_arr(i-1) then
			msgbox "有相同地點之資料存在！！",0,"同名警告"
			exit for
		end if
	next
end sub
sub thingcheck
	dim thing_arr(<%=thing_no%>)
	dim thing_no
	thing_no=<%=thing_no%>

<%	for i=1 to thing_no 
%>
		thing_arr(<%=i-1%>)="<%=thing_a(i-1)%>"
<%	next
%>
	for i=1 to thing_no
		thing_ck=Ucase(Trim(document.misc_add.keyword.value))
		if thing_ck=thing_arr(i-1) then
			msgbox "有相同事件之資料存在！！",0,"同名警告"
			exit for
		end if
	next
end sub
sub writercheck
	dim writer_arr(<%=writer_no%>)
	dim writer_no
	writer_no=<%=writer_no%>

<%	for i=1 to writer_no 
%>
		writer_arr(<%=i-1%>)="<%=writer_a(i-1)%>"
<%	next
%>
	for i=1 to writer_no
		writer_ck=Trim(document.misc_add.keyword.value)
		if writer_ck=writer_arr(i-1) then
			msgbox "有相同人員之資料存在！！",0,"同名警告"
			exit for
		end if
	next
end sub
</script>

<center>
<form name="misc_add" method=post action="misc_add_ok.asp">
<table border=1>
<tr width=720>
      <td width=720 colspan=5 align=center><font size=5 color="#0000ff">
      <b>選項資料新增畫面</b></font></td>
</tr>
<tr width=720>
	<td width=100 align=right><font size=3 color="#0000ff"><b>新增項目</b></font></td>
	<td width=100 align=left><font size=3 color="#0000ff">
		<select name="add_item" onchange="keycheck">
			<option value="0" >請選擇
			<option value="1" >人員
			<option value="2" >地點
			<option value="3" >事件
		</select>
	</font></td>
	<td width=100 align=right><font size=3 color="#0000ff"><b>新增文字</b></font></td>
	<td width=100 align=left><input type=text name="keyword" size="8" onblur="keycheck" onchange="keycheck"></td>
	<td width=320 align=left>&nbsp;&nbsp;
	<input type="submit" name="sent" value="新增資料">&nbsp;&nbsp;
	<input type="reset" name="reset" value="清除資料">
	</td>
</tr>
</table>

<table border=1>
	<tr width=720>

	<td width=180 align=center>現有人員項目<br>
		<select name="p_server" >
<%
	for i=1 to writer_no
		response.write "<option value='"&writer_a(i-1)&"'>"&writer_a(i-1)
	next
%>
		</select>	
	</td>	
	<td width=180 align=center>現有地點項目<br>
		<select name="p_place" >
<%
	for i=1 to place_no
		response.write "<option value='"&place_a(i-1)&"'>"&place_a(i-1)
	next
%>
		</select>	
	</td>
	<td width=180 align=center>現有事件項目<br>
		<select name="p_thing" >
<%
	for i=1 to thing_no
		response.write "<option value='"&thing_a(i-1)&"'>"&thing_a(i-1)
	next
%>
		</select>	

	</tr>
</table>
</form>
<td width=180 align=center><a href="./misc_edit.asp">回選項編修頁</a></td>
</center>
</body>
</html>