<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
	p_wk_content=trim(request("wk_content"))
	p_wk_item=trim(request("wk_item"))
	p_doing_date1=request("doing_date1")
	p_wk_class=request("wk_class")      '�u�@����

	p1_wk_exe=request("wk_exe")
	p_wk_att=request("wk_att")
	'p_wk_checker=request("wk_checker")     '�����H��
	p_wk_undoer=request("wk_undoer")     '�������H��
	p_wk_doer=request("wk_doer")       '���|�H��
	p_redo=request("redo")  '���s�q���ק�
	if p_redo="�O" then p_wk_item=p_wk_item&" [��"&date()&"�ק]"                   
if  instr(1,p_wk_doer,worker,1)=0 then p_wk_doer=p_wk_doer&","&worker	

p_wk_pjn=request("wk_pjn")          '�M�צW��

'if p_wk_pjn="0" or isnull(p_wk_pjn) then
'      p_pj_id=0
'      p_pj_02=null
'elseif  p_wk_pjn="" then
if trim(p_wk_pjn)="�A" then
      p_pj_id=null
      p_pj_02=null
else
      a_wk_pjn=split(p_wk_pjn,"�A",-1,1)
      p_pj_id=a_wk_pjn(0)
      p_pj_02=a_wk_pjn(1)
end if

p_wk_password=request("str_pwd")      '�[�K��r
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
	'----------2003/03/15�ץ� 
	'�ק��Ƥ�SQL���O�r�� �������
	'strSQL_edit="Update "&tb_name&" set wk_content='"&request("wk_content")&"'"
	'strSQL_edit=strSQL_edit & ",doing_date1=#"& request("doing_date1") &"#"
	'strSQL_edit=strSQL_edit & ",wk_item='"& request("wk_item") &"'"
	'strSQL_edit=strSQL_edit & " where wk_id =" & wk_id
	'����SQL���O
	'conDB.Execute strSQL_edit
	'---------------------------------------------------------
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id="&wk_id
rstObj1.open strSQL_show,conDB,1,3
rstobj1.MoveFirst
'Ū�����
po_wk_doer=rstObj1.fields("wk_doer")     '�ª��|�H��
po_wk_undoer=rstObj1.fields("wk_undoer")               '�¥������H��
po_checker=rstObj1.fields("wk_checker")                               '�¤w�����H��
po_group=rstObj1.fields("wk_group")    '�u�@�s��
'�ק���
rstObj1.fields("wk_content")= trim(p_wk_content)            '���e
rstObj1.fields("doing_date1")= p_doing_date1                   '������
rstObj1.fields("wk_item")= p_wk_item                                '�D��

if  po_group="�@��u�@" then
   rstObj1.fields("wk_class")= p_wk_class                               '����
else
   rstObj1.fields("pj_02")= p_pj_02                               '����
   rstObj1.fields("pj_id")= p_pj_id                               '����
end if

rstObj1.fields("wk_exe")= p1_wk_exe         '����H��
rstObj1.fields("wk_att")= p_wk_att         '�X�u�H��
rstObj1.fields("wk_doer")= p_wk_doer     '���|�H��

'�P�_�s���|�H��--------------------------------------------------------------------------------
pn_wk_doer=p_wk_doer        '�s���|�H��
pa_wk_doer=split(po_wk_doer,",",-1,1)
pa_wk_doer_no=ubound(pa_wk_doer)+1
for pai=1 to pa_wk_doer_no
   pn_wk_doer=replace(pn_wk_doer,pa_wk_doer(pai-1),"")
   pn_wk_doer=replace(pn_wk_doer,",,",",")
next
if left(pn_wk_doer,1)="," then pn_wk_doer=right(pn_wk_doer,len(pn_wk_doer)-1)
if right(pn_wk_doer,1)="," then pn_wk_doer=left(pn_wk_doer,len(pn_wk_doer)-1)

'�N�s���|�H���[�J�������H����-------------------------------------
if pn_wk_doer="" then
   pn_wk_undoer=po_wk_undoer
else
   if po_wk_undoer="" then
      pn_wk_undoer=pn_wk_doer
   else
      pn_wk_undoer=po_wk_undoer&","&pn_wk_doer
   end if
end if
rstObj1.fields("wk_undoer")=pn_wk_undoer
'�N�s���|�H���[�J�������H����-------------------------------------

if p_redo="�O" then
   rstObj1.fields("wk_undoer")=p_wk_doer
   rstObj1.fields("wk_checker")=""
else
	rstObj1.fields("wk_undoer")=p_wk_undoer
end if

rstObj1.fields("wk_password")=p_wk_password

rstObj1.UpdateBatch
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ�
set conDB=Nothing 

strURL1="wk_show.asp?wk_id="&wk_id
response.redirect(strURL1)
%>

</body>
</html>
