<% @codepage=950%>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
	pwd_headline=Request("pwd_headline")   '�K�X
%>

<%
if pwd_headline="3939" then
      '�N�u�@�C�����j�T��
      ' �s��Access��Ʈwdaywork.mdb
      DBpath=Server.MapPath("./database/daywork.mdb")
      strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
      '�إ߸�Ʈw�s������
      set conDB= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB.Open strCon
      '�}�Ҹ�ƪ�W��
      tb_name="work_data"
      '�إ߸�Ʈw�s������	
      set rstObj1=Server.CreateObject("ADODB.Recordset")
      strSQL_show="Select * from " & tb_name & " where wk_id="&wk_id
      rstObj1.open strSQL_show,conDB,1,3
      rstObj1.fields("headline")=5
      rstObj1.UpdateBatch
      '������ƶ�
      rstObj1.Close
      '���]����ܼ�
      set rstObj1=Nothing
      '������Ʈw
      conDB.Close
      '���]�����ܼ�
      set conDB=Nothing
    strURL1=session("hback_URL")
      'strURL1="wk_lst_doing.asp"
      response.redirect(strURL1)

else
    if isnull(pwd_headine) or pwd_headline="" then
         pwd_msg=""
    else
         pwd_msg="�K�X���~�I�I"
    end if
%>

<%
' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'Ū�����
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
wk_pjn=rstObj1.fields("pj_02")   '�M�צW��
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing
%>
<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<form name="form1" action="0_wk_headline_off.asp" method="post">
<input type="hidden" name="wk_id" value="<%=wk_id%>">
<input type="hidden" name="worker" value="<%=worker%>">
<font style="font-size:16pt;" color="red">�n�N�����j�T�������A�п�J�K�X�I�I</font><br>
<font style="font-size:12pt;" color="blue">�N�i�D���j�b�����]���O�������I�I</font><br>
�K�X�G<input type='password' name='pwd_headline' value='' style="width:100px;" > <br>
		<input type=submit name="editok" value="�T�w" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100px;">
		<input type=button name="goback1" value="�^�W�@��" onclick="history.back()" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100px;" >
<hr>
<%
if pwd_msg="�K�X���~�I�I" then
      response.write "<b>" & pwd_msg & "</b><hr>"
end if
%>

<table border=1 cellspacing=0 cellpadding=0>
<col width=120><col width=120><col width=120><col width=120><col width=120><col width=120>
<tr>
	<td align="center" colspan=2 rowspan=2><font size=4 color="red"><b>��ܳ�@�u�@��</b></font></td>
	<td align="right">�u�@�s�աG</td>
	<td>&nbsp;<%=wk_group%></td>
	<td align="right">�M�צW�١G</td>
	<td>&nbsp;<%=wk_pjn%></td>
</tr>
<tr>
	<td align="right">�u�@�s���G</td>
	<td>&nbsp;<%=wk_id%></td>
	<td align="right">�u�@�����G</td>
	<td>&nbsp;<%=wk_class%></td>
</tr>
<tr>
	<td align="right">���i�̡G</td>
	<td>&nbsp;<%=wk_order%></td>
	<td align="right">���i����G</td>
	<td>&nbsp;<%=undo_date1%></td>
	<td align="right">�������G</td>
	<td>&nbsp;<%=doing_date1%></td>
</tr>
<tr>
	<td align="right">
	���|�H���G
	</td>
	<td colspan=5>&nbsp;<%=wk_doer%></td>
</tr>
<tr>
	<td align="right">
	�����H���G
	</td>
	<td colspan=5>&nbsp;<%=wk_checker%></td>
</tr>
<tr>
	<td align="right">
	�������H���G
	</td>
	<td colspan=5>&nbsp;<%=wk_undoer%></td>
</tr>
<tr>
	<td align="right">
	�D���G
	</td>
	<td colspan=5>&nbsp;<%=wk_item%></td>
</tr>
<tr>
	<td align="right" valign="top">
	���椺�e�G
	</td>
	<td colspan=5>
	<%
	 wk_content=replace(wk_content,chr(13),"<br>")
	 response.write wk_content
	%>
	</td>
</tr>
</table>
</form>
<center>
</body>
</html>
<% end if %>