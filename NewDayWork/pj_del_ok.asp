<%@ Language=VBScript CODEPAGE=950 %>
<%
'=====================�򥻸��=============================		

'�M��id
	if request("p00")<>"" then 
		p_00=trim(request("p00"))
	else
		p_00=""
	end if
'�M�צW��
	'if request("p02")<>"" then 
		'p_02=trim(request("p02"))
	'else
		'p_02=""
	'end if

'�N�u�@�����M�צW�٧R�� 
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
	strSQL_show="Select * from " & tb_name & " where pj_id ="& p_00
	rstObj1.open strSQL_show,conDB,1,3
	totalput=rstObj1.recordcount
	if totalput=0 then	
	else
		rstobj1.MoveFirst
		for i=1 to totalput
			'Ū�����
			rstObj1.fields("pj_id")= null
			rstObj1.fields("pj_02")= ""
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

'�Nproject_data�����M�צW�٧R�� 
	' �s��Access��Ʈwdaywork.mdb
	DBpath=Server.MapPath("./database/daywork.mdb")
	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'�إ߸�Ʈw�s������
	set conDB= Server.CreateObject("ADODB.Connection")
	'�s����Ʈw	
	conDB.Open strCon
	'�}�Ҹ�ƪ�W��
	tb_name="project_data"
	'�R����Ƥ�SQL���O�r��
	strSQL_del="Delete from " & tb_name & " where pj_id =" & p_00
	'����SQL���O
	conDB.Execute strSQL_del
	'������Ʈw 
	conDB.Close
	'���]�����ܼ� 
	set conDB=Nothing 

str_url="pj_list.asp"
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
