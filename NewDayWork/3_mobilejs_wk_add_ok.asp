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
p_undo_date1=Request("undo_date1")         '���i���
p_doing_date1=Request("doing_date1")       '������
p_wk_class=Request("wk_class")                   '�u�@����
p_wk_group=Request("wk_group")                '�u�@�s��
p_wk_item=Request("wk_item")                     '�D��
p_wk_item=replace(p_wk_item,"'","��")
p_wk_content=Request("wk_content")         '���e
p_wk_content=replace(p_wk_content,"'","��")
p_wk_order=Request("wk_order")                 '���i��
p_all_worker=Request("all_worker")     '���|�H��
	p_wk_exe=request("wk_exe")       '����H��
if  instr(1,p_all_worker,worker,1)=0 then p_all_worker=p_all_worker&","&worker
p_all_worker=replace(p_all_worker," ","",1,-1,1)

'===========�P�_��ƬO�_��g����=================
str_error=""
if  p_doing_date1="" or isnull(p_doing_date1) or not(isdate(p_doing_date1)) then str_error=str_error&"[������]���~�C"
if  p_wk_item="" or isnull(p_wk_item) then str_error=str_error&"[�D��]�ťաC"
'if  p_wk_content="" or isnull(p_wk_content) then str_error=str_error&"[���e]�ťաC"
if  p_all_worker="" or isnull(p_all_worker) then str_error=str_error&"[���|�H��]�ťաC"
if  p_wk_order="" or isnull(p_wk_order) then response.redirect("./firstpage.asp")
if  p_wk_exe="" or isnull(p_wk_exe) then str_error=str_error&"[����H��]�ťաC"

if not(str_error="") then
   nexturl="3_mobilejs_wk_add.asp?ermsg="&str_error
   response.redirect(nexturl)
end if
'===========�P�_��ƬO�_��g����=================
p_wk_pjn=Request("wk_pjn")     '�M�צW��

if p_wk_pjn="0" or isnull(p_wk_pjn) then
      p_pj_id=null
      p_pj_02=null
else
      a_wk_pjn=split(p_wk_pjn,"�A",-1,1)
      p_pj_id=a_wk_pjn(0)
      p_pj_02=a_wk_pjn(1)
end if

if  instr(1,p_all_worker,worker,1)=0 then p_all_worker=p_all_worker&","&worker

p_wk_contenta=p_wk_content
'response.write "p_undo_date1=" & p_undo_date1 & "<br>"
'response.write "p_doing_date1=" & p_doing_date1 & "<br>"
'response.write "p_wk_class=" & p_wk_class & "<br>"
'response.write "p_wk_exe=" & p_wk_exe & "<br>"
'response.end
' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"

'�s�W��Ƥ�SQL���O�r��
strSQL_add="Insert into "&tb_name&" ("
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
if p_wk_pjn="0" or isnull(p_wk_pjn) then
else
   strSQL_add=strSQL_add & "pj_id,pj_02,"
end if
strSQL_add=strSQL_add & "wk_doer,wk_undoer) values ('"

strSQL_add=strSQL_add & p_undo_date1 &"','"
strSQL_add=strSQL_add & p_doing_date1 &"','"
strSQL_add=strSQL_add & p_wk_class &"','"
strSQL_add=strSQL_add & p_wk_group &"','"

strSQL_add=strSQL_add & p_wk_item&"','"
strSQL_add=strSQL_add & p_wk_contenta &"','"
strSQL_add=strSQL_add & p_wk_order &"','"
strSQL_add=strSQL_add & p_wk_exe &"','"
if p_wk_pjn="0" or isnull(p_wk_pjn) then
else
   strSQL_add=strSQL_add & p_pj_id &"','"
   strSQL_add=strSQL_add & p_pj_02 &"','"
end if

strSQL_add=strSQL_add & p_all_worker&"','"
strSQL_add=strSQL_add & p_all_worker&"')"

'����SQL���O
conDB.Execute strSQL_add

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " order by wk_id desc" 
rstObj1.open strSQL_show,conDB,1,1
rstObj1.movefirst
	p_tmp_id=rstObj1.fields("wk_id")
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
p_iptok=0
' �s��Access��Ʈwtemp-daywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"

'�s�W��Ƥ�SQL���O�r��
strSQL_add="Insert into "&tb_name&" (tmp_id,"
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
if p_wk_pjn="0" or isnull(p_wk_pjn) then
else
   strSQL_add=strSQL_add & "pj_id,pj_02,"
end if
strSQL_add=strSQL_add & "wk_doer,wk_undoer,ipt_ok) values ('"

strSQL_add=strSQL_add & p_tmp_id &"','"
strSQL_add=strSQL_add & p_undo_date1 &"','"
strSQL_add=strSQL_add & p_doing_date1 &"','"
strSQL_add=strSQL_add & p_wk_class &"','"
strSQL_add=strSQL_add & p_wk_group &"','"

strSQL_add=strSQL_add & p_wk_item&"','"
strSQL_add=strSQL_add & p_wk_contenta &"','"
strSQL_add=strSQL_add & p_wk_order &"','"
strSQL_add=strSQL_add & p_wk_exe &"','"
if p_wk_pjn="0" or isnull(p_wk_pjn) then
else
   strSQL_add=strSQL_add & p_pj_id &"','"
   strSQL_add=strSQL_add & p_pj_02 &"','"
end if

strSQL_add=strSQL_add & p_all_worker&"','"
strSQL_add=strSQL_add & p_all_worker&"',"
strSQL_add=strSQL_add & p_iptok&")"

'����SQL���O
conDB.Execute strSQL_add

'������Ʈw
conDB.Close
'���]�����ܼ�
set conDB=Nothing
'=============================================================================
   str_url="3_mobilejs_wk_show.asp?wk_id="&p_tmp_id
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
