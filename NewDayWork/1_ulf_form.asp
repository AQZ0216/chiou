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
wk_id=request("wk_id") 'Ū���u�@���ؤ�wk_id

if wk_id="" or isnull(wk_id ) then wk_id=0

%>

<HTML>
<HEAD>
<Title>�W���ɮ׵e��</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
<style type="text/css"><!--
body{font-family:'�з���';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<form id="form1" name="form1" method="post" action="1_ulf_form_ok.asp" enctype="multipart/form-data">
<input type="hidden" name="text" value="<%=wk_id%>" >
<table width=760 border=0 cellspacing=0 cellpadding=0 bgcolor="#FFFFBB">
<col width=60>
<col width=240>
<col width=60>
<col width=200>
<col width=80>
<tr>
<td colspan=5 align=center>
<b>�W�Ǥu�@���ت����ɮ�</b>
<a href="wk_show.asp?wk_id=<%=wk_id%>" title="�u�@wk_id=<%=wk_id%>">�^�u�@���e</a>
</td>
</tr>
<tr>
   <td align=right>�ɮ׻���</td>
   <td><input type="text" name="item" value="" style="width:100%" maxlength="40"></td>
   <td align=right>�ɮצW��</td>
   <td><input type="file" name="image" style="width:100%"></td>
   <td><input type="submit" name="button1" value="�W���ɮ�" style="width:100%"></td>
</tr>
<tr>
<td colspan=5 align=left style="padding-left:5px;">
�`�N�G�P�@�u�@�p�G�W�ǬۦP<font color=blue>�ɮצW��</font>�ɡA�N�|���N���ɮפλ����C
</td>
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
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id &" and del_ok = false"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
   response.write "�L����W�Ǫ���C"
else
%>
<table border=1 cellspacing=0 cellpadding=0 width=750 bgcolor="#CCEEFF">
<col width=40 style="text-align:center;">
<col width=340 style="padding-left:5px;text-align:left;">
<col width=260 style="padding-left:5px;text-align:left;">
<col width=100 style="text-align:center;">
<tr>
<td colspan=4>�{���W�Ǫ���C��</td>
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
<!--
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
</table>
-->
<hr>
</form>
</center>
</BODY>
</HTML>