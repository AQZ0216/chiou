<%@ Language=VBScript CODEPAGE=950 %>
<%
'����ˬd�u�@�O�_�s�b
function exist_wkid(pwk_id)
      ' �s��Access��Ʈwdaywork.mdb
      DBpath_fe=Server.MapPath("./database/daywork.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '�إ߸�Ʈw�s������
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB_fe.Open strCon_fe
      '�}�Ҹ�ƪ�W��
      tb_name_fe="work_data"
      '�إ߸�Ʈw�s������	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '������ƶ�
      rstObj1_fe.Close
      '���]����ܼ�
      set rstObj1_fe=Nothing
      '������Ʈw 
      conDB_fe.Close
      '���]�����ܼ�
      set conDB_fe=Nothing
      exist_wkid=totalput_fe
end function
%>
<HTML>
<HEAD>
<title>�Ҧ��W�Ǥ����[�ɮצC��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
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
strSQL_show="Select * from " & tb_name & " where del_ok = false order by wk_id desc, fl_date desc"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
else

%>
<table border=1 cellspacing=0 cellpadding=0 width=750 >
<col width=40 style="text-align:center;">
<col width=280 style="padding-left:5px;text-align:left;">
<col width=210 style="padding-left:5px;text-align:left;">
<col width=90 style="text-align:center;">
<col width=90 style="text-align:center;">
<tr>
<td colspan=5 style="font-size:15pt;color:blue;">�Ҧ����[�ɮצC��</td>
</tr>
<tr>
<td >�Ǹ�</td>
<td align=center >�ɮ׻���</td>
<td align=center >�ɮצW�� [�W�Ǫ�]</td>
<td >���ɤ��</td>
<td >�\��</td>
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
<td ><a href="./1_ulf_item_edit.asp?fl_id=<%=pfl_id%>" target="_self" title="�ק��ɮ׻����C" ><img src="./img/change.png" style="vertical-align:middle;height:16px;cursor:hand;border:0;" ></a>
<%=pfl_item%></td>
<td ><a href="./file_att/<%=pfl_name%>" target="_blank" title="<%=pfl_history%>"><%=str_pfl_name%></a> [<%=pfl_inputer%>]</td>
<td ><%=pfl_date%></td>
<td >
<% if exist_wkid(pwk_id)=1 then %>
<input type="button" name="shfile" value="�u"  onclick="file_sh('<%=pwk_id%>')" title="��ܭ�u�@���� [ wk_id=<%=pwk_id%> ] ���e�C">
<input type="button" name="addfile" value="�s"  onclick="file_add('<%=pwk_id%>')" title="�u�@���� [ wk_id=<%=pwk_id%> ] �s�W�ɮסC">
<% end if %>
<input type="button" name="delfile" value="�R"  onclick="file_del('<%=pfl_id%>')" title="�N�ɮקR���C">
<!-- <a href="1_ulf_form.asp?wk_id=<%=pwk_id%>" title="�s�W�ɮשΧ�s�ɮסC">�s</a>
<a href="1_ulf_del.asp?fl_id=<%=pfl_id%>" title="�R���ɮסC">�R</a> -->
</td>
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
</form>
</center>
<script language=vbscript>
sub file_sh(ppwk_id)
	ok=msgbox("�O�_�T�w�n��ܤu�@���ءH"&chr(13)&"wk_show.asp?wk_id="&ppwk_id,1,"�s�Wĵ�i")
	if ok=1 then 
		'location.href="wk_show.asp?wk_id="&ppwk_id
		window.open("wk_show.asp?wk_id="&ppwk_id)
	end if
end sub
sub file_add(ppwk_id)
	ok=msgbox("�O�_�T�w�n�s�W�ɮסH"&chr(13)&"1_ulf_form.asp?wk_id="&ppwk_id,1,"�s�Wĵ�i")
	if ok=1 then 
		location.href="1_ulf_form.asp?wk_id="&ppwk_id
	end if
end sub
sub file_del(ppfl_id)
	ok=msgbox("�O�_�T�w�n�R���ɮסH"&chr(13)&"1_ulf_del.asp?fl_id="&ppfl_id,1,"�R��ĵ�i")
	if ok=1 then 
		location.href="1_ulf_del.asp?fl_id="&ppfl_id
	end if
end sub

</script>
</body>
</html>