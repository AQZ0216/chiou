<%@ Language="VBScript" CODEPAGE=950 %>
<%
'�u�@�H���}�Cdaywork.mdb worker_data
dim worker_a()
' �s��Access��Ʈwdaywork.mdb
DBpath_a1=Server.MapPath("../holiday/database/crew.mdb")
strCon_a1="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_a1
'�إ߸�Ʈw�s������
set conDB_a1= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB_a1.Open strCon_a1
'�}�Ҹ�ƪ�W��
tb_name_a1="crew"
'�إ߸�Ʈw�s������	
set rstObj_a1=Server.CreateObject("ADODB.Recordset")
strSQL_a1="Select * from " & tb_name_a1 &" order by wk_sgp asc, wk_gp_sq, wkr_id asc"
rstObj_a1.open strSQL_a1,conDB_a1,3
'�p�����`��	
worker_no=rstObj_a1.recordcount
'���]�}�C�ƥ�
redim worker_a(Cint(worker_no))
rstObj_a1.MoveFirst
for i=1 to worker_no
	worker_a(i-1)=rstObj_a1.fields("worker")     '����W
'����U�@���O��
	rstObj_a1.MoveNext		
next
'������ƶ�
rstObj_a1.Close
'���]����ܼ� 
set rstObj_a1=Nothing
'������Ʈw 
conDB_a1.Close
'���]�����ܼ� 
set conDB_a1=Nothing 
%>

<%
'�קﬣ�u�H��
'p_order_old="�д@"
'p_order_new="Ellie"
%>
<html>
<head>
<title>����קﬣ�u�̦W��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body{font-family:'�з���';background-color:'##FFFFcc'}
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
����קﬣ�u��
<br>
	<input type="button" name="chk" value="�T�w�ק�" onclick="checka()" >
	<input type="reset" name="reset" value="�M�����" >
	<input type="button" name="giveup" value="�^�W�@��" onclick="history.back()" >
</tr>
<tr style="height:25px;" >
<td>�쬣�u��</td>
<td>
		<select name="p_order_old" style="width:100%">
		<option value="none" selected>�п�ܤH��...</option>
	<%
		for i=1 to worker_no
			response.write "<option value='"&worker_a(i-1)&"'>"&worker_a(i-1)
		next
	%>
		</select>
</td>
<td>�s���u��</td>
<td>
		<select name="p_order_new" style="width:100%">
		<option value="none" selected>�п�ܤH��...</option>
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
   if document.form1.p_order_old.value="none" then str_err="�п�ܭ쬣�u�̡I�I"
   if document.form1.p_order_new.value="none" then str_err=str_err & chr(13) &"�п�ܷs���u�̡I�I"
   if str_err="" then
      str_chk="�T�{�N�쬣�u�̡i"& document.form1.p_order_old.value & "�j" & chr(13)
      str_chk=str_chk &"�קאּ" & chr(13)
      str_chk=str_chk &"�s���u�̡i"& document.form1.p_order_new.value & "�j"  & chr(13)
      ok=msgbox(str_chk,64+1,"�T�{")
      if ok=1 then
         'msgbox ok,0,"���~ĵ�i"
   	  document.form1.submit
      end if
   else
	msgbox str_err,0,"���~ĵ�i"
   end if
End sub
-->
</script>
</body>
</html>
