<%
' �s��Access��Ʈw./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="linkdata"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " order by lk_row asc, lk_col asc"
rstObj1.open strSQL_show,conDB,3,1
'�p�����`��	
totalput=rstObj1.recordcount
if totalput=0 then
else
      '���ܲĤ@�����
      rstobj1.MoveFirst
      '�C�X��ƶ���
      for i=1 to totalput
      	'�]�w�ťո�Ƥ��ϬM
      p_id=rstObj1.fields("lk_id")		'id	
      p_01=rstObj1.fields("lk_url")		'�s�����}
      p_02=rstObj1.fields("lk_item")		'�u���D
      p_03=rstObj1.fields("lk_title")		'�y�z
      p_04=rstObj1.fields("lk_row")		'�C
      p_05=rstObj1.fields("lk_col")		'��
%>
<button class="w3-button w3-large w3-pale-yellow  w3-border w3-border-brown w3-round-large " style="margin:2px;padding:3px;width:150px;" onclick="url_new('<%=p_01%>')" >
<%=p_02%>
</button>
<%
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
	

