<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
'�P�_�O�_��J�u�@���� 
keyword=request("wk_class")
'if keyword="" then 
	'response.redirect("wk_add.asp")
'else
'end if

p_wk_id=Request("wk_id1")         'id
%>	
<html>
<head>
<title>��Ƨ���s�W</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body {  scrollbar-3dlight-color:#ffffff;
        scrollbar-arrow-color:#CCCCCC;
        scrollbar-base-color:#666633;
        scrollbar-darkshadow-color:#e6e6cc;
        scrollbar-face-color:#666666;
        scrollbar-highlight-color:#ffffff;
        scrollbar-shadow-color:#e6e6cc;
        scrollbar-track-color:#ffffff;
        margin:2mm 0mm 0mm 0mm;		/*��t�W�U���k*/
		font-family:'�з���';		/*�r��*/
		font-size:4.5mm; 			/*�r��j�p*/
		background-color:'#F0FFF0';
     }
td{
   margin:0 0 0 0;      /*��t�W�U���k*/
   border-color:'#808080'; /*���~���C��*/ 
   border-style:solid;     /*���~�ؽu��*/
   border-width:1px;    /*���~�ثp��*/  
   vertical-align:middle;  /*�r�髫������覡*/
   font-size:4.5mm;
   }
table{
   margin:0 0 0 0;      /*��t�W�U���k*/
   border-collapse:collapse;  /*��اΦ����X*/
   }
--></style>
</head>
<body>
<center>

<!-- �}�Ҹ�Ʈw -->
<%
p_doing_date1=Request("doing_date1")       '������
p_wk_item=Request("wk_item")                     '�D��
p_wk_item=replace(p_wk_item,"'","��")
p_wk_content=Request("wk_content")         '���e
p_wk_content=replace(p_wk_content,"'","��")
p_all_worker=Request("all_worker")     '���|�H��
p_wk_exe=request("wk_exe")       '����H��

if  instr(1,p_all_worker,worker,1)=0 then p_all_worker=p_all_worker&","&worker
p_all_worker=replace(p_all_worker," ","",1,-1,1)

p_wk_contenta=p_wk_content

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
strSQL_show="Select * from " & tb_name & " where wk_id="& p_wk_id &" order by wk_id desc" 
rstObj1.open strSQL_show,conDB,3,3

if rstObj1.recordcount=0 then
else	
	rstObj1.movefirst
		rstObj1.fields("doing_date1")=p_doing_date1	'������
		rstObj1.fields("wk_exe")=p_wk_exe						'����H��
		rstObj1.fields("wk_doer")=p_all_worker		'���|�H��
		rstObj1.fields("wk_content")=p_wk_content		'���e
		rstObj1.fields("wk_item")=p_wk_item					'�D��
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
'=============================================================================
'�Ȧs�ɮ�

' �s��Access��Ʈwtemp-daywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where (tmp_id ="& p_wk_id &" and ipt_ok=0) order by wk_id desc" 
rstObj1.open strSQL_show,conDB,3,3

if rstObj1.recordcount=0 then
else	
	rstObj1.movefirst
		rstObj1.fields("doing_date1")=p_doing_date1	'������
		rstObj1.fields("wk_exe")=p_wk_exe						'����H��
		rstObj1.fields("wk_doer")=p_all_worker		'���|�H��
		rstObj1.fields("wk_content")=p_wk_content		'���e
		rstObj1.fields("wk_item")=p_wk_item					'�D��
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
'=============================================================================
   str_url="3_mobilejs_wk_show.asp?wk_id="&p_wk_id
'   response.redirect(str_url)
'   str_url="wk_calendar_all.asp"
   response.redirect(str_url)

%>
<!-- <script language="Javascript">
	alert("��Ʒs�W�����I�I");
	location.href="wk_calendar_all.asp";
</script> -->

</body>
</html>
