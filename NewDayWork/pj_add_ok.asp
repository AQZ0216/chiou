<%@ Language=VBScript CODEPAGE=950 %>
<%
'=====================�򥻸��=============================		

'�u�@�s��p_00
	if request("p00")<>"" then 
		p_00=trim(request("p00"))
	else
		p_00=""
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

%>
<!-- Ū����� -->
<%
'�s�W��Ƥ�SQL���O�r��
strSQL_add="Insert into "&tb_name&" (pj_01,"						 
strSQL_add=strSQL_add & "pj_02,"				 
strSQL_add=strSQL_add & "pj_03) values ("						  
strSQL_add=strSQL_add &"'"&p_01&"',"	
strSQL_add=strSQL_add &"'"&p_02&"',"	
strSQL_add=strSQL_add &"'"&p_03&"')"
'����SQL���O
conDB.Execute strSQL_add

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " order by pj_id desc"
rstObj1.open strSQL_show,conDB,3,3
'�p�����`��	
totalput=rstObj1.recordcount
'���ܲĤ@����� 
rstobj1.MoveFirst
newid=rstObj1.fields("pj_id")
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 

if p_00="" then
else
	'�N�W�٥[�J�u�@�� 
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
	strSQL_show="Select * from " & tb_name & " where wk_id="& p_00 
	rstObj1.open strSQL_show,conDB,1,3
	rstobj1.MoveFirst
	'Ū�����
	rstObj1.fields("pj_id")= newid
	rstObj1.fields("pj_02")= trim(p_02)
	rstObj1.UpdateBatch
	'������ƶ�
	rstObj1.Close
	'���]����ܼ� 
	set rstObj1=Nothing
	'������Ʈw 
	conDB.Close
	'���]�����ܼ� 
	set conDB=Nothing 
end if
str_url="pj_show.asp?p_id="&newid
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
