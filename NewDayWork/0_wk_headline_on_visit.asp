<% @codepage=950%>
<%
	'Ū���H���m�W
	worker = Session("worker")
%>

<%
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
      strSQL_show="Select * from " & tb_name & " where wk_item like '%�ӳX%' or wk_item like '%���X%' or wk_item like '%�줽�q%'"
      rstObj1.open strSQL_show,conDB,1,3
	'�p�����`��	
	totalput=rstObj1.recordcount
	if totalput= 0 then
	else
		'���ܲĤ@�����
		rstObj1.MoveFirst
	    for kj=1 to totalput      
      	rstObj1.fields("headline")=10
      	rstObj1.UpdateBatch
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
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
	�w�N���X��ƦC�����j�T���C
</center>
</body>
</html>
