<%@ Language="VBScript" CODEPAGE=950 %>
<%
'工作人員陣列daywork.mdb worker_data
dim worker_a()
' 連結Access資料庫daywork.mdb
DBpath_a1=Server.MapPath("../holiday/database/crew.mdb")
strCon_a1="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_a1
'建立資料庫連結物件
set conDB_a1= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB_a1.Open strCon_a1
'開啟資料表名稱
tb_name_a1="crew"
'建立資料庫存取物件	
set rstObj_a1=Server.CreateObject("ADODB.Recordset")
strSQL_a1="Select * from " & tb_name_a1 &" order by wk_sgp asc, wk_gp_sq, wkr_id asc"
rstObj_a1.open strSQL_a1,conDB_a1,3
'計算資料總數	
worker_no=rstObj_a1.recordcount
'重設陣列數目
redim worker_a(Cint(worker_no))
rstObj_a1.MoveFirst
for i=1 to worker_no
	worker_a(i-1)=rstObj_a1.fields("worker")     '中文名
'移到下一筆記錄
	rstObj_a1.MoveNext		
next
'關閉資料集
rstObj_a1.Close
'重設資料變數 
set rstObj_a1=Nothing
'關閉資料庫 
conDB_a1.Close
'重設物件變數 
set conDB_a1=Nothing 
%>

<%
'修改派工人員
'p_order_old="羽婷"
'p_order_new="Ellie"
%>
<html>
<head>
<title>整體修改派工者名稱</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body{font-family:'標楷體';background-color:'##FFFFcc'}
-->
</style>
</head>
<body>
<center>
<form name="form1" method="post" action="wk_revise_order_ok.asp" >
<table border=1 cellspacing=0 cellpadding=0 >
<col style="width:100px;padding:2 2 0 2;" align=center>
<col style="width:200px;padding:2 2 0 2;" align=center>
<col style="width:100px;padding:2 2 0 2;" align=center>
<col style="width:200px;padding:2 2 0 2;" align=center>

<tr style="height:30px;" bgcolor="#f0f0ff">
<td colspan=4 align=center>
整體修改派工者
<br>
	<input type="button" name="chk" value="確定修改" onclick="checka()" >
	<input type="reset" name="reset" value="清除資料" >
	<input type="button" name="giveup" value="回上一頁" onclick="history.back()" >
</tr>
<tr style="height:25px;" >
<td>原派工者</td>
<td>
		<select name="p_order_old" style="width:100%">
		<option value="none" selected>請選擇人員...</option>
	<%
		for i=1 to worker_no
			response.write "<option value='"&worker_a(i-1)&"'>"&worker_a(i-1)
		next
	%>
		</select>
</td>
<td>新派工者</td>
<td>
		<select name="p_order_new" style="width:100%">
		<option value="none" selected>請選擇人員...</option>
	<%
		for i=1 to worker_no
			response.write "<option value='"&worker_a(i-1)&"'>"&worker_a(i-1)
		next
	%>
		</select>
</td>
</tr>
</table>
</form>
<script Language="VBScript">
<!--
Sub checka()
   str_err=""
   if document.form1.p_order_old.value="none" then str_err="請選擇原派工者！！"
   if document.form1.p_order_new.value="none" then str_err=str_err & chr(13) &"請選擇新派工者！！"
   if str_err="" then
      str_chk="確認將原派工者【"& document.form1.p_order_old.value & "】" & chr(13)
      str_chk=str_chk &"修改為" & chr(13)
      str_chk=str_chk &"新派工者【"& document.form1.p_order_new.value & "】"  & chr(13)
      ok=msgbox(str_chk,64+1,"確認")
      if ok=1 then
         'msgbox ok,0,"錯誤警告"
   	  document.form1.submit
      end if
   else
	msgbox str_err,0,"錯誤警告"
   end if
End sub
-->
</script>
</body>
</html>
