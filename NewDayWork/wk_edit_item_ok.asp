<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	p_wk_item_old=trim(request("wk_item_old"))
	p_wk_item_new=trim(request("wk_item_new"))
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
strSQL_show="Select * from " & tb_name & " where wk_item like '"& p_wk_item_old &"' order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,1,3
totalput=rstObj1.recordcount
rstobj1.MoveFirst
if totalput=0 then
else
      '�ק���
      for kj=1 to totalput
            rstObj1.fields("wk_item")= p_wk_item_new                                '�D��
      	'����U�@���O��
      		rstObj1.MoveNext
      		if rstObj1.EOF=True then exit for
      next
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

strURL1="wk_query_oki.asp?q_text="&p_wk_item_new
response.redirect(strURL1)
%>

</body>
</html>
