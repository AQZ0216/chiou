<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
%>
<%
'BASP21.DLL�N�ɮפW���{��
'�Х�����w��RegSvr32 Basp21.dll
'�i�N��椤��text�]��X��i�}�C�A����response.write �ܼơA�N�i�Hprint�X�ӤF
'-------------------------------------------------------------------
'�W�Ǫ����ɮ׵e��
fl_id=request("fl_id") 'Ū���ɮ�id

if fl_id="" or isnull(fl_id) then
   myURL="1_ulf_list.asp"
   Response.Redirect (myURL)
end if

'���[�ɮצC��
' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/attach_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="file_data"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where fl_id =" & fl_id &" and del_ok = false"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
else
	'�C�X��ƶ���
	rstobj1.MoveFirst
	for fi=1 to totalput
	'Ū�����
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")
		pfl_item=rstObj1.fields("fl_item")
		pfl_date=rstObj1.fields("fl_date")
		str_none=pwk_id&"_"
		str_pfl_name=right(pfl_name,len(pfl_name)-len(pwk_id)-1)
	'����U�@���O��
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
end if
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
<Title>�ק�����ɮפ�����</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
<style type="text/css"><!--
body{font-family:'�з���';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<form id="form1" name="form1" method="post" action="1_ulf_item_edit_ok.asp" >
<input type="hidden" name="fl_id" value="<%=fl_id%>" >
<table width=760 border=0 cellspacing=0 cellpadding=0>
<col width=70>
<col width=240>
<col width=70>
<col width=190>
<col width=60>
<tr>
<td colspan=5 align=center>
<b>�ק�����ɮ׻���</b>
<a href="wk_show.asp?wk_id=<%=pwk_id%>" title="��ܤu�@���e">�u�@wk_id=<%=pwk_id%></a>
</td>
</tr>
<tr>
   <td align=right>�ɮ׻����G</td>
   <td><input type="text" name="item" value="<%=pfl_item%>" style="width:100%" maxlength="40"></td>
   <td align=right>�ɮצW�١G</td>
   <td><%=pfl_name%></td>
   <td><input type="submit" name="button1" value="�T�w�ק�" style="width:100%"></td>
</tr>
</table>

<%
'���[�ɮצC��
' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/attach_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="file_data"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & pwk_id &" and del_ok = false"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
else
%>
<table border=1 cellspacing=0 cellpadding=0 width=750 >
<col width=40 style="text-align:center;">
<col width=340 style="padding-left:5px;text-align:left;">
<col width=260 style="padding-left:5px;text-align:left;">
<col width=100 style="text-align:center;">
<tr>
<td colspan=4>�{������C��</td>
</tr>
<tr>
<td >�Ǹ�</td>
<td align=center >�ɮ׻���</td>
<td align=center >�ɮצW�� [�W�Ǫ�]</td>
<td >���ɤ��</td>
</tr>
<%
	'�C�X��ƶ���
	rstobj1.MoveFirst
	for fi=1 to totalput
	'Ū�����
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
	'����U�@���O��
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>

</table>
<%
end if
'������ƶ�
rstObj1.Close
'���]����ܼ�
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ�
set conDB=Nothing
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
strSQL_show="Select * from " & tb_name & " where wk_id =" & pwk_id
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
%>
<%
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
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
	<td align="center" colspan=2 rowspan=2><font size=4 color="red"><b>��ܳ�@�u�@��</b></font></td>
	<td align="right">�u�@�s�աG</td>
	<td><%=showspace(wk_group)%>
	</td>
	<td align="right">�M�צW�١G</td>
	<td><%=showspace(wk_pjn)%>
	</td>
</tr>

<tr>
	<td align="right">�u�@�s���G</td>
	<td><%=showspace(wk_id)%>
	</td>
	<td align="right">�u�@�����G</td>
	<td><%=showspace(wk_class)%>
	</td>
</tr>

<tr>
	<td align="right">���i�̡G</td>
	<td><%=showspace(wk_order)%>
	</td>
	<td align="right">���i����G</td>
	<td><%=showspace(undo_date1)%>
	</td>
	<td align="right">�������G</td>
	<td><%=showspace(doing_date1)%>
	</td>
</tr>
<tr>
	<td align="right">
	���|�H���G
	</td>
	<td colspan=5><%=showspace(wk_doer)%>
	</td>
</tr>
<tr>
	<td align="right">
	�����H���G
	</td>
	<td colspan=5><%=showspace(wk_checker)%>
	</td>
</tr>
<tr>
	<td align="right">
	�������H���G
	</td>
	<td colspan=5><%=showspace(wk_undoer)%>
	</td>
</tr>
<tr>
	<td align="right">
	�D���G
	</td>
	<td colspan=5><%=showspace(wk_item)%>
	</td>
</tr>
<tr>
	<td align="right" valign="top">
	���椺�e�G
	</td>
	<td colspan=5>
	<%
	if wk_content="" or isnull(wk_content) then
	  wk_content_a=wk_content
	else
	  wk_content_a=replace(wk_content,chr(13),"<br>")
	end if
	response.write  wk_content_a
	%>
<!-- 	<TEXTAREA name="wk_content" rows="10" style="width:100%;" readonly><%=wk_content%></TEXTAREA>
 -->
 	</td>
</tr>
</table>
<hr>
</form>
</center>
</BODY>
</HTML>