
<html>
<head>
<title>選項資料刪除</title>
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</head>
<body >
<!-- #Include file = "./array_place_misc.inc" -->	
<!-- #Include file = "./array_thing_misc.inc" -->	
<!-- #Include file = "./array_writer_misc.inc" -->	
<script language=vbscript>
sub writercheck
	document.misc_add.add_item.value=1
	document.misc_add.keyword.value=document.misc_add.p_server.value
end sub
sub placecheck
	document.misc_add.add_item.value=2
	document.misc_add.keyword.value=document.misc_add.p_place.value
end sub
sub thingcheck
	document.misc_add.add_item.value=3
	document.misc_add.keyword.value=document.misc_add.p_thing.value
end sub

</script>

<center>
<form name="misc_add" method=post action="misc_del_ok.asp">
<table border=1>
<tr width=720>
      <td width=720 colspan=5 align=center><font size=5 color="#0000ff">
      <b>選項資料刪除畫面</b></font></td>
</tr>
<tr width=720>
	<td width=100 align=right><font size=3 color="#0000ff"><b>刪除項目</b></font></td>
	<td width=100 align=left><font size=3 color="#0000ff">
		<select name="add_item">
			<option value="0" >請選擇
			<option value="1" >人員
			<option value="2" >地點
			<option value="3" >事件
		</select>
	</font></td>
	<td width=100 align=right><font size=3 color="#0000ff"><b>刪除文字</b></font></td>
	<td width=100 align=left><input type=text name="keyword" size="8" onblur="keycheck" onchange="keycheck"></td>
	<td width=320 align=left>&nbsp;&nbsp;
	<input type="submit" name="sent" value="刪除資料">&nbsp;&nbsp;
	<input type="reset" name="reset" value="清除資料">
	</td>
</tr>
</table>

<table border=1>
	<tr width=720>
	</td>
	<td width=180 align=center>現有人員項目<br>
		<select name="p_server" onchange="writercheck">
<%
	for i=1 to writer_no
		response.write "<option value='"&writer_a(i-1)&"'>"&writer_a(i-1)
	next
%>
		</select>	
	</td>	
	<td width=180 align=center>現有地點項目<br>
		<select name="p_place" onchange="placecheck">
<%
	for i=1 to place_no
		response.write "<option value='"&place_a(i-1)&"'>"&place_a(i-1)
	next
%>
		</select>	
	</td>
	<td width=180 align=center>現有事件項目<br>
		<select name="p_thing" onchange="thingcheck">
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