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

'�P�_�O�_�O���Ʃʤ��i
j_date=Request("doing_date1")
repeat_type=request("repeat_type")
end_date=Request("end_date")
'repeat_no=request("repeat_no")
'repeat_unit=request("repeat_unit")
'select case repeat_unit
'   case "u_week"
'      end_date= dateadd("ww",repeat_no,j_date)
'   case "u_month"
'      end_date= dateadd("m",repeat_no,j_date)
'   case "u_year"
'      end_date= dateadd("yyyy",repeat_no,j_date)
'end select

'response.write "����=" & end_date
'response.end
 str_date=j_date
 pre_date=j_date
 ntk=1
' pre_yy=year(pre_date)
' pre_mm=month(pre_date)
 'pre_dd=day(pre_date)
 check_s=false
select case repeat_type
   case "cs_1"          '�榸
     repeat_num=1
   case "cs_week1"   '�C�P�@��
      do
         next_date=dateadd("ww",1,pre_date)
         if next_date >= cdate(end_date) then
            check_s=true
            'exit do
         else
            str_date=str_date&","&next_date
            pre_date= next_date
         end if
      loop until check_s=true
   case "cs_week2"   '��P�@��
      do
         next_date=dateadd("ww",2,pre_date)
         if next_date >= cdate(end_date) then
            check_s=true
            'exit do
         else
            str_date=str_date&","&next_date
            pre_date= next_date
         end if
      loop until check_s=true
   case  "cs_month1"      '�C��@��
      do
         next_date=dateadd("m",ntk,pre_date)
         if next_date >= cdate(end_date) then
            check_s=true 
         else
            str_date=str_date&","&next_date
            'pre_date= next_date
            ntk=ntk+1
         end if
      loop until check_s=true
   case  "cs_year1"     '�C�~�@��
       do
         next_date=dateadd("yyyy",1,pre_date)
         if next_date >= cdate(end_date) then
            check_s=true
         else
            str_date=str_date&","&next_date
            pre_date= next_date
         end if
      loop until check_s=true
   case "cs_m_first_monday"     '�C��Ĥ@�ӬP���@
       do
         next_date=dateadd("m",1,pre_date)
         next_date=dateserial(year(next_date),month(next_date),1)
         kww=Weekday(next_date)
         kadd=((7+2)-kww) mod 7
         next_date=dateadd("d",kadd,next_date)
         if next_date >= cdate(end_date) then
            check_s=true
         else
            str_date=str_date&","&next_date
            pre_date= next_date
         end if
      loop until check_s=true
   case "cs_manual"     '�ۭq���    2011/11/02�s�W
            str_date01=request("rp_dates")    '�ۭq���
            str_date01=replace(str_date01,chr(13),",")
            str_date01=replace(str_date01,",,",",")
            if right(str_date01,1)="," then str_date01=left(str_date01,len(str_date01)-1)
            str_date=trim(str_date01)
   case else
end select

'response.write " str_date=" &  str_date
'response.end

'����}�C
date_arr=Split(str_date, ",", -1, 1)
date_num=ubound(date_arr)+1

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
p_headline=Request("headline1")
p_undo_date1=Request("undo_date1")
'p_doing_date1=Request("doing_date1")
p_wk_class=Request("wk_class")
p_wk_group=Request("wk_group")
p_wk_item=Request("wk_item")
p_wk_item=replace(p_wk_item,"'","��")
p_wk_content=Request("wk_content")
p_wk_content=replace(p_wk_content,"'","��")
p_wk_order=Request("wk_order")
p_wk_exe=Request("wk_exe")
p_wk_att=Request("wk_att")
p_all_worker=Request("all_worker")     '���|�H��

p_wk_pjn=Request("wk_pjn")     '�M�צW��

if p_wk_pjn="0" or isnull(p_wk_pjn) then
      p_pj_id=null
      p_pj_02=null
else
      a_wk_pjn=split(p_wk_pjn,"�A",-1,1)
      p_pj_id=a_wk_pjn(0)
      p_pj_02=a_wk_pjn(1)
end if
p_str_pwd=Request("str_pwd")  '�[�K��r
'========='�P�_�O�_�q���y��==================
golf_ok=request("golf_ok")
if golf_ok="�O" then
   p_wk_contenta=now()&"���i���y���C"& chr(13) &p_wk_content
else
   p_wk_contenta=p_wk_content
end if
'========='�P�_�O�_�q���y��==================

if  instr(1,p_all_worker,worker,1)=0 then p_all_worker=p_all_worker&","&worker

' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"
%>

<%
for zki=1 to date_num      '==========�j��}�l================

   if isdate(date_arr(zki-1))=true then '====='�P�_����榡�O�_���T=========================

      '�s�W��Ƥ�SQL���O�r��
      strSQL_add="Insert into "&tb_name&" ("
      strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,headline,"
      strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,wk_att,"
      if p_wk_pjn="0" or isnull(p_wk_pjn) then
      else
         strSQL_add=strSQL_add & "pj_id,pj_02,"
      end if

      strSQL_add=strSQL_add & "wk_password,wk_doer,wk_undoer) values ('"

      strSQL_add=strSQL_add & p_undo_date1 &"','"
      strSQL_add=strSQL_add & date_arr(zki-1) &"','"
      strSQL_add=strSQL_add & p_wk_class &"','"
      strSQL_add=strSQL_add & p_wk_group &"','"
      strSQL_add=strSQL_add & p_headline &"','"

      strSQL_add=strSQL_add & p_wk_item&"','"
      strSQL_add=strSQL_add & p_wk_contenta &"','"
      strSQL_add=strSQL_add & p_wk_order &"','"
      strSQL_add=strSQL_add & p_wk_exe &"','"
      strSQL_add=strSQL_add & p_wk_att &"','"
      if p_wk_pjn="0" or isnull(p_wk_pjn) then
      else
         strSQL_add=strSQL_add & p_pj_id &"','"
         strSQL_add=strSQL_add & p_pj_02 &"','"
      end if

      strSQL_add=strSQL_add & p_str_pwd&"','"
      strSQL_add=strSQL_add & p_all_worker&"','"
      strSQL_add=strSQL_add & p_all_worker&"')"

      '����SQL���O
      conDB.Execute strSQL_add
   end if                                        '====='�P�_����榡�O�_���T=========================

next                    '==========�j�鵲��================

'������Ʈw
conDB.Close
'���]�����ܼ�
set conDB=Nothing

'========='�P�_�O�_�q���y��==================
'golf_ok=request("golf_ok")

if golf_ok="�O" then
   p_udate=p_undo_date1    '���i���
   p_ddate=date_arr(0)          '������
   p_witem=p_wk_item            '�D��
   p_wcontent=p_wk_content    '���e
'   str_golfurl="http://114.32.81.117:90/wk_add_c_ok.asp?undo_date1="&p_udate&"&doing_date1="&p_ddate&"&wk_item="&p_witem&"&wk_content="&p_wcontent
   'str_golfurl="http://192.168.0.125/chiou/daywork/wk_add_c_ok.asp?undo_date1="&p_udate&"&doing_date1="&p_ddate&"&wk_item="&p_witem&"&wk_content="&p_wcontent
   'window.open(str_golfurl)
   p_remark="���i�̡G"&p_wk_order&"(��j�a��)�C"& now() &"���i�C"
%>
<!-- <script language="Javascript">
	window.open("http://192.168.0.125/chiou/daywork/wk_add_c_ok.asp?undo_date1=<%=p_udate%>&doing_date1=<%=p_ddate%>&wk_item=<%=p_witem%>&wk_content=<%=p_wcontent%>");
</script> -->
<form name="form1" action="http://1.34.48.166:90/wk_add_c_ok.asp" method="post" >
<!--	<form name="form1" action="http://114.32.81.117:90/wk_add_c_ok.asp" method="post" >-->
<!-- <form name="form1" action="http://192.168.0.125/chiou/daywork/wk_add_c_ok.asp" method="post" > -->
<font style="text-align:center;font-size:5mm;color:blue;">���q���i�w��Ʒs�W�����I�I</font>
<hr>
<table border=0>
<col style="width:120px;text-align:right;">
<col style="width:500px;">
<td colspan=2 style="text-align:center;font-size:5mm;color:blue;">
���I��i�T�w�j�A�H�T�w�N��Ƶo�G�찪���Ҳy���C
</td>
<tr>
<td>���i����G</td>
<td><input type="text" name="undo_date1" value="<%=p_udate%>" readonly></td>
</tr>
<tr>
<td>�������G</td>
<td><input type="text" name="doing_date1" value="<%=p_ddate%>" readonly></td>
</tr><tr>
<td>�D���G</td>
<td><input type="text" name="wk_item" value="<%=p_witem%>" style="width:100%;" ></td>
</tr><tr>
<td>���椺�e�G</td>
<td><TEXTAREA name="wk_content" rows="5" style="width:100%;" ><%=p_remark%>&#013;<%=p_wcontent%></TEXTAREA></td>
</tr>
<tr>
<td colspan=2 style="text-align:center;">
<input type="submit" name="bt" value="�i�T�w�s�W���i���j�����Ҿǭb�j" style="text-align:center;font-size:5mm;color:blue;">
</td>
</tr>
 </table>
<!-- <br>
<a href="http://192.168.0.125/chiou/daywork/wk_add_c_ok.asp?undo_date1=<%=p_udate%>&doing_date1=<%=p_ddate%>&wk_item=<%=p_witem%>&wk_content=<%=p_wcontent%>" target="_blank" style="color:blue;background-color:#DDDDDD;">�i�T�w�j</a> <br>
 -->
<hr>
<!-- <a href="wk_calendar_all.asp" target="_self">�^����</a><br> -->
</form>
<!-- <script language="Javascript">
	alert("��Ʒs�W�����I�I");
	location.href="wk_calendar_all.asp";
</script> -->

<%
         'response.write "��Ʒs�W����"
else
'   str_url="wk_calendar_all.asp"
'   response.redirect(str_url)
   strbackURL=session("strbackURL")
'   if strbackURL="" or isnull(strbackURL) or not( strbackURL="wk_lst_doing.asp" ) or not( left(strbackURL,19)="wk_calendar_all.asp" ) then strbackURL="wk_calendar_all.asp"
   if strbackURL="" or isnull(strbackURL) then
      'response.write  "strbackURL=null<br>"
      strbackURL="wk_Calendar_all.asp"
   else
      if strbackURL="wk_lst_doing.asp" or left(strbackURL,19)="wk_Calendar_all.asp" then
         'response.write  "strbackURL=wk_lst_doing.asp <br>"
      else
         'response.write  "strbackURL="& strbackURL &"<br>"
         'response.write  "not strbackURL=wk_lst_doing.asp <br>"
         strbackURL="wk_Calendar_all.asp"
      end if
   end if
   'response.write  strbackURL
   response.redirect(strbackURL)

end if

    'response.end
'
'response.write "��Ʒs�W����"
'response.end
'if golf_ok="�O" then response.end
'str_url="wk_calendar_all.asp"
'response.redirect(str_url)
%>
<!-- <script language="Javascript">
	alert("��Ʒs�W�����I�I");
	location.href="wk_calendar_all.asp";
</script> -->

</body>
</html>
