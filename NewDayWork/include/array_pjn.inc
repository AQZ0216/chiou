<%
'�u�@�H���}�Cdaywork.mdb worker_data
dim pjn_a()
dim pjnid_a()
%>
<%
' �s��Access��Ʈwdaywork.mdb
DBpath_a1=Server.MapPath("./database/daywork.mdb")
strCon_a1="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_a1
'�إ߸�Ʈw�s������
set conDB_a1= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB_a1.Open strCon_a1
'�}�Ҹ�ƪ��W��
tb_name_a1="project_data"
'�إ߸�Ʈw�s������	
set rstObj_a1=Server.CreateObject("ADODB.Recordset")
strSQL_a1="Select * from " & tb_name_a1
rstObj_a1.open strSQL_a1,conDB_a1,1,1
'�p�����`��	
pjn_no=rstObj_a1.recordcount
if pjn_no=0 then
else
      '���]�}�C�ƥ�
      redim pjn_a(Cint(pjn_no))
      redim pjnid_a(Cint(pjn_no))
      rstObj_a1.MoveFirst
      for i=1 to pjn_no
      	pjn_a(i-1)=rstObj_a1.fields("pj_02")
      	pjnid_a(i-1)=rstObj_a1.fields("pj_id")
      '����U�@���O��		
      	rstObj_a1.MoveNext		
      next
end if
'������ƶ�
rstObj_a1.Close
'���]����ܼ� 
set rstObj_a1=Nothing
'������Ʈw 
conDB_a1.Close
'���]�����ܼ�
set conDB_a1=Nothing 
%>


	