<%@ Language=VBScript CODEPAGE=950 %>
<%
'=====================�򥻸��=============================		

'�u�@�s��p_00
	if request("p00")<>"" then 
		p_00=trim(request("p00"))
	else
		p_00=""
	end if
'�M��id p_id
	if request("p_id")<>"" then 
		p_id=trim(request("p_id"))
	else
		p_id=null
	end if
'�M�׽s��
'	if request("p01")<>"" then 
'		p_01=trim(request("p01"))
'	else
'		p_01=""
'	end if
'�M�צW��
'	if request("p02")<>"" then 
'		p_02=trim(request("p02"))
'	else
'		p_02=""
'	end if
'�M�׻���p_03
'	if request("p03")<>"" then 
'		p_03=trim(request("p03"))
'	else
'		p_03=""
'	end if

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
	p_idold=rstObj1.fields("pj_id")
	rstObj1.fields("pj_id")= p_id
	rstObj1.fields("pj_02")= ""
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
if isnull(p_id) then
	str_url="pj_show.asp?p_id="&p_idold
	response.redirect(str_url) 
else
	str_url="pj_show.asp?p_id="&p_id
	response.redirect(str_url) 
end if

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
