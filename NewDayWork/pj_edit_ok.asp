<%@ Language=VBScript CODEPAGE=950 %>
<%
'=====================�򥻸��=============================		
'�M��id
	if request("p_id")<>"" then 
		pj_id=request("p_id")
	else
		pj_id=""
	end if

'�M�׽s��
	if request("p01")<>"" then 
		p_01=trim(request("p01"))
	else
		p_01=""
	end if
'�M�צW��
	if request("p02")<>"" then 
		p_02=trim(request("p02"))
	else
		p_02=""
	end if
'�M�׻���p_03
	if request("p03")<>"" then 
		p_03=trim(request("p03"))
	else
		p_03=""
	end if

%> 
<!-- �}�Ҹ�Ʈw -->
<%
' �s��Access��Ʈw./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="project_data"
	'�إ߸�Ʈw�s������	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " where pj_id="& pj_id 
	rstObj1.open strSQL_show,conDB,1,3
	rstobj1.MoveFirst
	'Ū�����
	rstObj1.fields("pj_01")= trim(p_01)
	rstObj1.fields("pj_02")= trim(p_02)
	rstObj1.fields("pj_03")= trim(p_03)
	rstObj1.UpdateBatch
	'������ƶ�
	rstObj1.Close
	'���]����ܼ� 
	set rstObj1=Nothing

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 

	'�N�u�@��pj_02�M�צW�٧� 
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
	strSQL_show="Select * from " & tb_name & " where pj_id="& pj_id 
	rstObj1.open strSQL_show,conDB,1,3
	'�p�����`��	
	totalput=rstObj1.recordcount
	if totalput=0 then
	else
		'�C�X��ƶ���
		rstobj1.MoveFirst
		for j=1 to totalput
			rstObj1.fields("pj_02")= trim(p_02)'�M�צW�٧�
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

str_url="pj_show.asp?p_id="&pj_id
response.redirect(str_url) 

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�s�W</title>
<style type="text/css"><!--
body{font-family:'�s�ө���';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>

</body>
</html>
