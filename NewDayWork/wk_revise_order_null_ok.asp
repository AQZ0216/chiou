<%@ Language=VBScript CODEPAGE=950 %>
<%
'�קﬣ�u�H��
'p_order_old="�д@"
'p_order_new="Ellie"
'p_order_old=request("p_order_old")
'p_order_new=request("p_order_new")
'�ק�ťլ��u�̤����

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
strSQL_show="Select * from " & tb_name & " where wk_order like '' or isnull(wk_order) order by wk_id asc"
rstObj1.open strSQL_show,conDB,3,3
'�p�����`��
totalput=rstObj1.recordcount
if totalput=0 then
else
   rstObj1.MoveFirst
   for j=1 to totalput
      '�ק���
      p_doer=rstObj1.fields("wk_doer")            '�u�@�H��
      if instr(1,p_doer,",",1)=0 then
         p_order= p_doer            '���u��
      else
         if instr(1,p_doer,"���z",1)>0 then
            p_order="���z"
         elseif instr(1,p_doer,"Ellie",1)>0 then
    	       p_order="Ellie"
         elseif instr(1,p_doer,"����",1)>0 then
    	       p_order="����"
         elseif instr(1,p_doer,"���",1)>0 then
    	       p_order="���"
         elseif instr(1,p_doer,"���`",1)>0 then
    	       p_order="���`"
         end if
      end if
         rstObj1.fields("wk_order")=p_order
         response.write j &".���u�̧אּ�i"& p_order &"�j�C<br>"

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

</body>
</html>
