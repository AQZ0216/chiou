<% @codepage=950%>

<%
pyymmdd=year(date())& right("0"& month(date()),2) & right("0"& day(date()),2)
pfilename="calendar_google_"&pyymmdd&".csv"
Response.AddHeader "Content-Disposition","attachment;filename="&pfilename
Response.ContentType = "application/vnd.ms-csv"
%>
<%
'�u�@id p_wkid
p_wkid=request("p_wkid")
'if p_wkid="" or isnull(p_wkid) then p_wkid=""
'response.write "p_wkid="& p_wkid &"<br>"
arr_wkid=split(p_wkid,",",-1,1)
no_wkid=ubound(arr_wkid)+1
'response.write 	p_wkid
'response.end
%>
<%
'��X��ƪ����ରutf-8 65001
Response.Charset="utf-8"
Session.Codepage=65001
%>
<%
	str_00="Subject"'���ʦW�� (���n)�C
	str_01="Start Date"'���ʪ��Ĥ@�� (���n)�C
	str_02="Start Time"'���ʶ}�l�ɶ��C
	str_03="End Date"'���ʪ��̫�@�ѡC
	str_04="End Time"'���ʵ����ɶ��C
	str_05="All Day Event"'�o�Ӭ��ʬO�_�����Ѭ��ʡC�p�G�O���Ѭ��ʡA�п�J True�F�_�h�п�J False�C
	str_06="Description"'���ʻ����Ϊ����C
	str_07="Location"'���ʦa�I�C
	str_08="Private"'�o�Ӭ��ʬO�_���p�H���ʡC�p�G�O�p�H���ʡA�п�J True�F�_�h�п�J False�C
	Response.Write str_00 & "," & str_01 & "," & str_02 & "," & str_03 & "," & str_04 & "," & str_05 & "," & str_06 & "," & str_07 & "," & str_08 & vbCrLf
%>
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

for kj=1 to no_wkid
	ppwkid=arr_wkid(kj-1)
	'-----------------------------------
	'�إ߸�Ʈw�s������	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name &" where wk_id ="& ppwkid &""
	'Response.Write querystr & vbCrLf
	rstObj1.open strSQL_show,conDB,1,1
	totalput=rstObj1.recordcount
	if totalput=0 then
	else
		'�C�X��ƶ���
		rstobj1.MoveFirst
		'Ū�����
			doing_date1=rstObj1.fields("doing_date1")'���u���
			wk_item=replace(trim(rstObj1.fields("wk_item")),",","�A")'�D��
			wk_item=replace(wk_item,";",":")'�D��
			wk_content=left(rstObj1.fields("wk_content"),200)'�u�@���e���O
			wk_content=replace(wk_content,",","�A")'�u�@���e���O
			wk_content=replace(wk_content,chr(13),"�C")'�u�@���e���O
			wk_content=replace(wk_content,chr(10),"")'�u�@���e���O
			str1_02a=left(wk_item,5)
			if not(isnumeric(left(str1_02a,2))) then
				str1_02a="08:00"
			end if
			str1_04a=Mid(wk_item,7,5)
			if not(isnumeric(left(str1_04a,2))) then
				str1_04a=str1_02a
			end if
		str1_00=wk_item	'���ʦW�� (���n)�C
		str1_01=doing_date1		'���ʪ��Ĥ@�� (���n)�C
		str1_02=str1_02a	'���ʶ}�l�ɶ��C
		str1_03=doing_date1		'���ʪ��̫�@�ѡC
		str1_04=str1_04a		'���ʵ����ɶ��C
		str1_05="False"				'�o�Ӭ��ʬO�_�����Ѭ��ʡC�p�G�O���Ѭ��ʡA�п�J True�F�_�h�п�J False�C
		str1_06=wk_content	'���ʻ����Ϊ����C
		str1_07="taipei"'���ʦa�I�C
		str1_08="False"'�o�Ӭ��ʬO�_���p�H���ʡC�p�G�O�p�H���ʡA�п�J True�F�_�h�п�J False�C
		Response.Write str1_00 & "," & str1_01 & "," & str1_02 & "," & str1_03 & "," & str1_04 & "," & str1_05 & "," & str1_06 & "," & str1_07 & "," & str1_08 & vbCrLf
	end if
	'������ƶ�
	rstObj1.Close
	'���]����ܼ� 
	set rstObj1=Nothing
	'-----------------------------------
next

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
<%
'��X��ƪ����ରutf-8 65001
Response.Charset="big-5"
Session.Codepage=950
%>