<% @codepage=950%>

<%
pyymmdd=year(date())& right("0"& month(date()),2) & right("0"& day(date()),2)
pfilename="calendar_google_"&pyymmdd&".csv"
Response.AddHeader "Content-Disposition","attachment;filename="&pfilename
Response.ContentType = "application/vnd.ms-csv"
%>
<%
'Ū���d�߱����� 
querystr=" where "
querystra=""
querystrb="�d�߱���G"

'�D��p_wk_item
p_wk_item=request("p_wk_item")
if p_wk_item="" or p_wk_item="����" then
	p_wk_item="����"
else
	querystra=querystra & "wk_item like '%"& p_wk_item &"%' and "
	querystrb=querystrb & "[�D��="& trim(p_wk_item) &"]"
	querystrc=querystrc & "p_wk_item"& trim(p_wk_item) &"&"
end if

'����H��p_wk_exe
p_wk_exe=trim(request("p_wk_exe"))
'p_wk_exe="���F"	
if p_wk_exe="" or p_mtclass="����" then
	p_wk_exe="����"
else
	querystra=querystra & "(wk_exe like '%"& p_wk_exe &"%' or wk_exe like '����H��' ) and "
	querystrb=querystrb & "[����H��="&trim(p_wk_exe)&"]"
	querystrc=querystrc & "p_wk_exe="&trim(p_wk_exe)&"&"
end if


'������p_doing_date1a
p_doing_date1a=trim(request("p_doing_date1a"))	
'p_doing_date1a="2016/3/1"
if p_doing_date1a="" or p_doing_date1a="����" then
	p_doing_date1a="����"
else
	querystra=querystra & "(doing_date1 >= #"& p_doing_date1a &"# ) and "
	querystrb=querystrb & "[������="&trim(p_doing_date1a)&"]"
	querystrc=querystrc & "p_doing_date1a="&trim(p_doing_date1a)&"&"
end if

'������p_doing_date1b
p_doing_date1b=trim(request("p_doing_date1b"))	
'p_doing_date1b="2016/4/1"
if p_doing_date1b="" or p_doing_date1b="����" then
	p_doing_date1b="����"
else
	querystra=querystra & "(doing_date1 <= #"& p_doing_date1b &"# ) and "
	querystrb=querystrb & "[������="&trim(p_doing_date1b)&"]"
	querystrc=querystrc & "p_doing_date1b="&trim(p_doing_date1b)&"&"
end if

	querystr=querystr & querystra
	len_a=len(querystr)
	if len_a=7 then querystr=" "
      if trim(querystr)="where" then querystr=" "
	if right(querystr,4)="and " then querystr=left(querystr,len_a-4)
	len_c=len(querystrc)
	if right(querystrc,1)="&" then querystrc=left(querystrc,len_c-1)
	
%>
<%
'��X��ƪ����ରutf-8 65001
Response.Charset="utf-8"
Session.Codepage=65001
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

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'strSQL_show="Select * from " & tb_name & " where wk_class like '%"&wk_class&"%' and wk_doer like '%"& wk_man &"%' order by doing_date1 desc"
strSQL_show="Select * from " & tb_name & querystr &" order by doing_date1 asc"
'Response.Write querystr & vbCrLf
rstObj1.open strSQL_show,conDB,1,1
totalput=rstObj1.recordcount
%>

<%
if totalput=0 then
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
else
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

	'�C�X��ƶ���
	rstobj1.MoveFirst
	for i=1 to totalput
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

	'����U�@���O��
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	

end if
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
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