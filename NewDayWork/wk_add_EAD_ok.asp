<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<%
'ead�|ĳ���i
	'Ū���H���m�W
	worker = Session("worker")
'��J���
j_date=request("p_ymd")

'�P�_�O�_�O���Ʃʤ��i
end_date=dateadd("m",1,j_date)
 str_date=""
 pre_date=dateadd("d",-1,j_date)   '�}�l���
 ntk=1
'�C��P�@�ܶg��
      do
         next_date=dateadd("d",ntk,pre_date)
         if next_date >= cdate(end_date) then
            check_s=true 
         else
            if Weekday(next_date) > 1 and Weekday(next_date) < 7 then
               str_date=str_date&","&next_date
            end if
            ntk=ntk+1
         end if
      loop until check_s=true

'����}�C
if left(str_date,1)="," then str_date=right(str_date,len(str_date)-1)
date_arr=Split(str_date, ",", -1, 1)
date_num=ubound(date_arr)+1
%>	
<%
'�Ҧ��H���r��
     str_allworker=worker_a(0)
	for i=2 to worker_no
		str_allworker=str_allworker&","& worker_a(i-1)
	next
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
'
 pct_date=date()   '���i���
 pdt_date=date()   '������
'���i���
p_undo_date1=pct_date
'������
'p_doing_date1= pdt_date
'�u�@����
p_wk_class=""
'�u�@�s��
p_wk_group="�@��u�@"
'�D��
'p_wk_item="08:20-09:00 EAD�|ĳ"
'p_wk_item="08:45-09:15 EAD�|ĳ"	'2014/11/24 �}�l
p_wk_item="08:30-09:0 E0AD�|ĳ"	'2014/12/16 �}�l
'���椺�e
'p_wk_content="�C��08:20-09:00 EAD�|ĳ"&chr(13)&"/�|ĳ�����A�Фť��Z"
'p_wk_content="�C��08:45-09:15 EAD�|ĳ"&chr(13)&"�|ĳ�����A�Фť��Z"		'2014/11/24 �}�l
p_wk_content="�C��08:30-09:00 EAD�|ĳ"&chr(13)&"�|ĳ�����A�Фť��Z"		'2014/12/16 �}�l
'���i��
p_wk_order=worker

p_wk_exe="���`,���,���"
'���|�H��
p_all_worker=str_allworker     '���|�H��

'response.write "���i���p_undo_date1="&p_undo_date1&"�C<br>"
'response.write "������p_doing_date1="&str_date&"�C<br>"
'response.write "�u�@�s��p_wk_group="&p_wk_group&"�C<br>"
'response.write "�u�@�D��p_wk_item="&p_wk_item&"�C<br>"
'response.write "���椺�ep_wk_content="&p_wk_content&"�C<br>"
'response.write "���i��p_wk_order="&p_wk_order&"�C<br>"
'response.write "���|�H��p_all_worker="&p_all_worker&"�C<br>"
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

for zki=1 to date_num


'�s�W��Ƥ�SQL���O�r��
strSQL_add="Insert into "&tb_name&" ("
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
strSQL_add=strSQL_add & "wk_doer,wk_undoer) values ('"

strSQL_add=strSQL_add & p_undo_date1 &"','"
strSQL_add=strSQL_add & date_arr(zki-1) &"','"
strSQL_add=strSQL_add & p_wk_class &"','"
strSQL_add=strSQL_add & p_wk_group &"','"

strSQL_add=strSQL_add & p_wk_item&"','"
strSQL_add=strSQL_add & p_wk_content &"','"
strSQL_add=strSQL_add & p_wk_order &"','"
strSQL_add=strSQL_add & p_wk_exe &"','"
strSQL_add=strSQL_add & p_all_worker&"','"
strSQL_add=strSQL_add & p_all_worker&"')"

'����SQL���O
conDB.Execute strSQL_add

next

'������Ʈw
conDB.Close
'���]�����ܼ�
set conDB=Nothing
   str_url="wk_calendar_all.asp?nMonth="& month(j_date) &"&nYear="& year(j_date)
   response.redirect(str_url)
%>
</body>
</html>
