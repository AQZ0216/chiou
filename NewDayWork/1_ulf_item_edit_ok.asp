<%@ Language=VBScript CODEPAGE=950 %>

<%
	'Ū���H���m�W
	worker = Session("worker")
	fl_id=Request("fl_id")
	pfl_item=Request("item")
	p_fl_history=now()&"�e"&worker&"�f�ק��ɮ׻����C"
%>
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
strSQL_show="Select * from " & tb_name & " where fl_id =" & fl_id &" and del_ok = false"
rstObj1.open strSQL_show,conDB,3,3
totalput=rstObj1.recordcount
if totalput=0 then
else
	'�C�X��ƶ���
	rstobj1.MoveFirst
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")      '�ɮצW��
		pfl_date=rstObj1.fields("fl_date")           '���ɤ��
		rstObj1.fields("fl_item")=pfl_item           '�ɮ׻���
		rstObj1.fields("fl_history")=rstObj1.fields("fl_history") & chr(13) & p_fl_history '�ק�L�{
         rstObj1.UpdateBatch
end if
'������ƶ�
rstObj1.Close
'���]����ܼ�
set rstObj1=Nothing
'������Ʈw
conDB.Close
'���]�����ܼ�
set conDB=Nothing

'response.write "�ɮקR������"
myURL="wk_show.asp?wk_id="&pwk_id
Response.Redirect (myURL)
%>

<HTML>
<HEAD>
<Title>�W���ɮץ\��{��</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
</HEAD>
<BODY>
<center>
�ɮקR������!!
<hr>
<a href="wk_show.asp?wk_id=<%=pwk_id%>" target="_self">�^�u�@����</a>
</center>
</BODY>
</HTML>