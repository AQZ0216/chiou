<%@ Language=VBScript CODEPAGE=950 %>
<%
'�קﬣ�u�H��
'p_order_old="�д@"
'p_order_new="Ellie"
p_order_old=request("p_order_old")
p_order_new=request("p_order_new")

%>
<html>
<head>
<title>��ƭק�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>
<!-- �}�Ҹ�Ʈw -->
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
strSQL_show="Select * from " & tb_name & " where wk_order like'%" & p_order_old &"%' order by wk_id asc"
rstObj1.open strSQL_show,conDB,3,3
'�p�����`��
totalput=rstObj1.recordcount
if totalput=0 then
else
   rstObj1.MoveFirst
   for j=1 to totalput
      '�ק���
      rstObj1.fields("wk_order")= trim(p_order_new)            '���u��
	'����U�@���O��
	rstObj1.MoveNext
	if rstObj1.EOF=True then exit for
   next
end if
rstObj1.UpdateBatch
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ�
set conDB=Nothing 
%>

��ƭק粒���C
�Ҧ����u�̸�ơi<%=p_order_old%>�j�אּ�i<%=p_order_new%>�j�C
</body>
</html>
