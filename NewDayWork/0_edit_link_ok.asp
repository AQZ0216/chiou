<%@ Language=VBScript CODEPAGE=950 %>
<%

p_id = request("p_id")   'id
'�]�wŪ����ƽs��
p_01=Request("p_01")    '�s�����}
p_02=Request("p_02")    '²�u���D
p_03=Request("p_03")    '�y�z

'Ū���������O�}�C
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
   strSQL_show="Select * from " & tb_name & " where lk_id="& p_id &" order by lk_id asc"
rstObj1.open strSQL_show,conDB,3,3
'�p�����`��	
totalput01=rstObj1.recordcount
'�C�X��ƶ���

      rstObj1.fields("lk_url")	=p_01	'�s�����}
      rstObj1.fields("lk_item")=p_02		'�u���D
      rstObj1.fields("lk_title")=p_03		'�y�z

rstObj1.UpdateBatch
'������ƶ�
rstObj1.Close
'���]����ܼ�
set rstObj1=Nothing
'������Ʈw
conDB.Close
'���]�����ܼ�
set conDB=Nothing

   str_url="firstpage_elink.asp"   'Ū���C����}
   response.redirect(str_url)


%>