<%@ Language=VBScript CODEPAGE=950 %>

<%
   'Ū���H���m�W
   worker = Session("worker")

%>
<%
'�𰲸��
function hd_man(p_hdate)
   pstr_hdman =""
    ' �s��Access��Ʈwholiday.mdb
    DBpath_fh=Server.MapPath("../holiday/database/holiday.mdb")
    strCon_fh="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fh
    '�إ߸�Ʈw�s������
    set conDB_fh= Server.CreateObject("ADODB.Connection")
    '�s����Ʈw	
    conDB_fh.Open strCon_fh
    '�}�Ҹ�ƪ�W��
    tb_name_fh="�𰲩���"
	'�إ߸�Ʈw�s������
	set rstObj1_fh=Server.CreateObject("ADODB.Recordset")
	strSQL_show_fh="Select * from " & tb_name_fh & " where �𰲤� = #"& p_hdate &"# order by ���Oid asc "
	rstObj1_fh.open strSQL_show_fh,conDB_fh,3,1
	totalput_fh=rstObj1_fh.recordcount
if not rstObj1_fh.EOF then
	rstObj1_fh.Movefirst
	for i = 1 to totalput_fh
		hd_id=rstObj1_fh.fields("hd_id")
		icon_id=rstObj1_fh.fields("���Oid")
		hd_hrs=rstObj1_fh.fields("�𰲮ɼ�")
		hd_check=rstObj1_fh.fields("�T�{")
		hd_man=rstObj1_fh.fields("���u�m�W")'���u�m�W
		hd_img=left(rstObj1_fh.fields("���O�W��"),1)
		hd_cname=right(rstObj1_fh.fields("���O�W��"),len(rstObj1_fh.fields("���O�W��"))-1)
		'�M�w���O�C��
		select case icon_id
'		   Case 1  f_color = "#FF0000"    '���G����C
'		   Case 2  f_color = "#000000"    '���G�ư��C
'		   Case 3  f_color = "#0000FF"    '��G�f���C
'		   Case 4  f_color = "#BB5500"    '���G�����C
'		   Case 5  f_color = "#000000"    '���G�ల�C
'		   Case 6  f_color = "#00FF00"    '���G�~���C
'		   Case 7  f_color = "#FF0088"    '���G�S��C
'		   Case 8  f_color = "#EE7700"    '���G�����C
'		   Case 9  f_color = "#BBBB00"    '���G�B���C
'		   Case 15  f_color = "#000000"   '���G�����d�C
'		   Case 16  f_color = "#FF0000"   '���G�ƯZ�C
'		   Case Else   f_color = "#FF0000"
		   Case 1  f_color = "#000000"    '���G����C
		   Case 2  f_color = "#000000"    '���G�ư��C
		   Case 3  f_color = "#000000"    '��G�f���C
		   Case 4  f_color = "#000000"    '���G�����C
		   Case 5  f_color = "#000000"    '���G�ల�C
		   Case 6  f_color = "#000000"    '���G�~���C
		   Case 7  f_color = "#000000"    '���G�S��C
		   Case 8  f_color = "#000000"    '���G�����C
		   Case 9  f_color = "#000000"    '���G�B���C
		   Case 15  f_color = "#000000"   '���G�����d�C
		   Case 16  f_color = "#000000"   '���G�ƯZ�C
		   Case Else   f_color = "#000000"
		End Select
		if icon_id=1 or icon_id=15 then
		    if icon_id=15 then
		       pstr_hdman = pstr_hdman & "<font style='font-size:15px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man & "</font><br>"
   	           end if
		else
		  pstr_hdman = pstr_hdman & "<font style='font-size:15px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man &"("& hd_hrs&")&nbsp;</font><br>"
		end if
		    'Response.Write "</font><br>"
		rstObj1_fh.MoveNext
		if rstObj1_fh.EOF=true then exit for
	next
else
end if
	'������ƶ�
	rstObj1_fh.Close
	'���]����ܼ� 
	set rstObj1_fh=Nothing
    '������Ʈw
    conDB_fh.Close
    '���]�����ܼ� 
    set conDB_fh=Nothing
  hd_man=pstr_hdman
end function
%>
<%
'�d�߬O�_������
Function exist_attach(pwk_id)
      ' �s��Access��Ʈwdaywork.mdb
      DBpath_fe=Server.MapPath("./database/attach_file.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '�إ߸�Ʈw�s������
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB_fe.Open strCon_fe
      '�}�Ҹ�ƪ�W��
      tb_name_fe="file_data"
      '�إ߸�Ʈw�s������	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where del_ok = false and wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '������ƶ�
      rstObj1_fe.Close
      '���]����ܼ�
      set rstObj1_fe=Nothing
      '������Ʈw 
      conDB_fe.Close
      '���]�����ܼ�
      set conDB_fe=Nothing
      exist_attach=totalput_fe
End Function

%>
<%
'�]�w�ܼ� 
dim dbConn, rs, nDex, nMonth, nYear, dtDate

' Get the current date
dtDate = Now()

' Set the Month and Year
   nWeeks=request("nWeeks")             '�P��Ū���ϥΪ�
   nYear=request("nYear")             '�~��Ū���ϥΪ�
if nWeeks = "" then nWeeks =DatePart("ww",Date()) 'Ū���g��
if nYear = "" then nYear = Year(Date())
	nMonth=month( Now())   
   '�]�w�g�Ƥ��Ĥ@��wkn_1 �γ̫�@��wkn_2
   date00=dateserial(nYear,1,1)  '�~�����Ĥ@��
  dayweek_2f=DateAdd("d",7-DatePart("w",date00)+1,date00)   '�ĤG�g���Ĥ@��

   wkn_1=DateAdd("d",(nWeeks-2)*7 ,dayweek_2f)
   wkn_2=DateAdd("d",((nWeeks-2)*7+6) ,dayweek_2f)
   wkn_0=DateAdd("d",-7 ,wkn_1)


'nMonth = Request("nMonth")
'nYear = Request("nYear")
'if nMonth = "" then nMonth = Month(dtDate)
'if nYear = "" then nYear = Year(dtDate)

select case cint(Weekday(date()))
case 1
   cswday="�P����"
case 2
   cswday="�P���@"
case 3
   cswday="�P���G"
case 4
   cswday="�P���T"
case 5
   cswday="�P���|"
case 6
   cswday="�P����"
case 7
   cswday="�P����"

end select

'�]�w�P���C�� 
bgc1="#F0FFF0"    '�H����lightyellow 
bgc6="#F0FFF0" 'lightskyblue
bgc7="#F0FFF0" 'lightgreen

' Set the date to the first of the current month
'dtDate = DateSerial(nYear, nMonth, 1)
dtDate=wkn_1

if int(nMonth)<10 then
   strnMonth="0"&cstr(nMonth)
else
   strnMonth=cstr(nMonth)
end if
dcodeym=cstr(nYear)&strnMonth

'�]�wsession("strbackURL")
'strbackURL="wk_week_now.asp?nWeeks="&nWeeks&"&nYear="&nYear
'session("strbackURL")=strbackURL

%>
<%
'�]�w�W�@�� 
if nMonth = 1 then 
   pre1month=12
   pre1year=nYear-1
else
   pre1month=nMonth-1
   pre1year=nYear
end if
pre2month=nMonth
pre2year=nYear-1
pre3month=nMonth
pre3year=nYear+1
if nMonth = 12 then 
   pre4month=1
   pre4year=nYear+1
else
   pre4month=nMonth+1
   pre4year=nYear
end if
%>
<%
pstart=dateserial(nYear,nMonth,1)
pend=dateadd("m",1,pstart)
'�d�ߥ���O�_�إ�EAD�|ĳ         p_wk_item="08:20-09:00 EAD�|ĳ"
function find_ead(pstart,pend)
      ' �s��Access��Ʈwdaywork.mdb
      DBpath_ead=Server.MapPath("./database/daywork.mdb")
      strCon_ead="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_ead
      '�إ߸�Ʈw�s������
      set conDB_ead= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw   
      conDB_ead.Open strCon_ead
      '�}�Ҹ�ƪ�W��
      tb_name_ead="work_data"
      '�إ߸�Ʈw�s������	
      set rstObj1_ead=Server.CreateObject("ADODB.Recordset")
      strSQL_show_ead="Select * from " & tb_name_ead & " where wk_item like '08:20-09:00 EAD�|ĳ' and wk_order like '���z' and doing_date1 >= #"& pstart &"# and doing_date1 < #"& pend &"# order by doing_date1 asc"
      rstObj1_ead.open strSQL_show_ead,conDB_ead,3,1
      totalput_ead=rstObj1_ead.recordcount
      '������ƶ�
      rstObj1_ead.Close
      '���]����ܼ� 
      set rstObj1_ead=Nothing
      '������Ʈw 
      conDB_ead.Close
      '���]�����ܼ� 
      set conDB_ead=Nothing
      find_ead=totalput_ead
end function
pchk_ead=find_ead(pstart,pend)

%>

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
%>

<HTML>
<HEAD>
<title>�˪O���D</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{
   margin:10px 0 0 0;      /*��t�W�U���k*/
   font-family:'�s�ө���';      /*�r��*/
   font-size:12.5px;          /*�r��j�p*/
   background-color:'#F0FFF0'; /*�I���C��*/
   /*letter-spacing:2px;  */
   }
notetext{
   font-family:'�s�ө���';      /*�r��*/
   font-size:12.5px;          /*�r��j�p*/
   }
daytext{
   font-family:'�s�ө���';      /*�r��*/
   font-size:12.5px;          /*�r��j�p*/
   }

/*�s��LINK������*/
A:link{color:black}     /*���s�����r���C��*/
A:visited{color:black}  /*���g�s�����r���C��*/
A:active{color:black}   /*�s�����r���C��*/
--></style>

</HEAD>
<body>
<center>
<form method="post" name="form1" action="">
<table border=0 bgcolor="gray" cellpadding=1 style="width:100%;border-width:0px;width:970px;">
<col style="width:56%;background-color:#F0FFF0;">
<col style="width:44%;background-color:#F0FFF0;">
<tr style="height:25px">
   <td align=center style="font-size:18px;letter-spacing:1px;" >
      <b><%=worker%>�������u�@����</b>
   <td colspan=1 align=center style="font-size:15px;cursor:hand;">
   ���ѬO&nbsp;�褸<%=cstr(Year(date()))%>�~<%=cstr(Month(date()))%>��<%=cstr(Day(date()))%>��&nbsp;<%=cstr(cswday)%> 
   <br>
   ��ܦ~���G<%=nYear%>
   ��ܶg�ơG<%=nWeeks%>
</table>
   <table border=0 bgcolor="gray" cellpadding=1 style="width:100%;border-width:0px;width:970px;">
   <col style="width:14%;background-color:<%=bgc7%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc6%>;">
   <tr style="height:45px;">
      <td align=center><font color="black"><b>�P����<br>Sunday</b></font></td>
      <td align=center><font color="black"><b>�P���@<br>Monday</b></font></td>
      <td align=center><font color="black"><b>�P���G<br>Tuesday</b></font></td>
      <td align=center><font color="black"><b>�P���T<br>Wednesday</b></font></td>
      <td align=center><font color="black"><b>�P���|<br>Thursday</b></font></td>
      <td align=center><font color="black"><b>�P����<br>Friday</b></font></td>
      <td align=center><font color="black"><b>�P����<br>Saturday</b></font></td></tr>
   <tr bgcolor="#ffffc0" style="height:60px;">
      <% 
      ' Add blank cells until the proper day�W�[�Ů����w��m 
      for nDex = 1 to Weekday(dtDate) - 1
         Response.Write "<td bgcolor=""#c0c0c0"">&nbsp;</td>"
      next
      '�}�l��J��� 
      do
         '����ഫ 
         if int(Day(dtDate))<10 then
            strnDay="0"&cstr(Day(dtDate))
         else
            strnDay=cstr(Day(dtDate))
         end if
         dcodeymd=dcodeym&strnDay
         dcodeymd_a=dtDate
         '�]�w�C�� 
         select case cint(Weekday(dtDate))
         case 1
            bgc=bgc7
         case 7
            bgc=bgc6
         case else 
            bgc=bgc1 
         end select

         '����椺��g��r 
         if int(nYear)=int(Year(date()))and int(nMonth)=int(Month(date()))and int(Day(dtDate))=int(day(date())) then
            Response.Write "<td valign=""top"" bgcolor=""#f0ffae"" >"
         else        
            Response.Write "<td valign=""top"" bgcolor="&bgc&" >"
         end if
               Response.Write "<table border=0 cellpadding=0 style=""width:100%;height:100%;border-width:0px;"">"
               Response.Write "<tr style=""width:100%;height:10px;"">"
               Response.Write "<td align=center style=""width:50%;background-color:#dcdcdc;font-size:13px;"">"
               Response.Write Day(dtDate)
               Response.Write "<td align=center style=""width:50%;background-color:#dcdcdc;"">"
               Response.Write "<font style=""font-size:13px;""><a href='wk_add.asp?datecode="&dtDate&"'>�s�W"
               Response.Write "</a></font>"
               'Response.Write "<font style=""text-size:3mm;""><b>&nbsp;"
               'Response.Write "</b></font>"
               Response.Write "<tr><td colspan=2 align=left valign=top>"
         '��J�ƥ��� 
            'Response.Write "<font size=""-1""><b>" &dcodeymd&"</b></font><br>"
            'Response.Write "<font size=""-1""><b>" &totalput&"</b></font><br>"
            '�إ߸�Ʈw�s������  
            set rstObj1=Server.CreateObject("ADODB.Recordset")
            strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and wk_undoer like '%"&worker&"%' order by wk_item asc , wk_id asc"
            rstObj1.open strSQL_show,conDB,3,1
            totalput=rstObj1.recordcount
         if not rstObj1.EOF then
            rstObj1.Movefirst
            for i = 1 to totalput
            	 wk_headline=rstObj1.fields("headline")  '�]���O
               wk_id=rstObj1.fields("wk_id")
               '�ˬd�O�_������ exist_attach(wk_id)
               attach_no=exist_attach(wk_id)
               if attach_no=0 then
                  str_colors="color:#000000;"
               else
                  str_colors="color:#0000FF;"
               end if
               if rstObj1.fields("wk_password")="" or isnull(rstObj1.fields("wk_password")) then
               else
                  str_colors="color:#0000FF;"
               end if
								p_nexe=rstObj1.fields("wk_exe")	'����H��
								if Instr(1, p_nexe, worker, 1)>0 or Instr(1, p_nexe, "����", 1)>0 then
									str_bgc="background-color:#99FF99;"	
								else
									str_bgc=""
								end if
               Response.Write "<font style='font-size:13px;"& str_bgc & str_colors &"' >" & i & "�B<a href='wk_show.asp?wk_id="&wk_id&"' style='letter-spacing:1.5pt;font-size:11pt;"& str_colors &"' >" & replace (rstObj1.fields("wk_item"),"���׶�","<font color=fuchsia >���׶�</font>")
               Response.Write "</a></font>"
if wk_headline>5 then response.write "<img src='./img/gnome_chess.png' title='�]���O�T��' width=19 style='vertical-align:top;'>"                
               Response.Write "<br>"
               rstObj1.MoveNext
               if rstObj1.EOF=true then exit for
            next
         else
         end if
            '������ƶ�
            rstObj1.Close
            '���]����ܼ� 
            set rstObj1=Nothing
             str_hdman=hd_man(dcodeymd_a)     '�𰲤H�����
            response.write str_hdman
            Response.Write "</table>"
            Response.Write "</td>"
                  
            if WeekDay(dtDate) = 7 then 
               Response.Write "</tr>" & vbCrLf & "<tr  bgcolor=""#ffffc0"" style=""height:60px;"">"
            end if
            dtDate = DateAdd("d", 1, dtDate)
            
      loop until ( dtDate > wkn_2 )
'      loop until (Month(dtDate) <> CInt(nMonth))
      '�����g���� 
         ' Add blank cells to fill out the rest of the month if needed�W�[�Ů����w��m
         if Weekday(dtDate) <> 1 then 
            for nDex = Weekday(dtDate) to 7
               Response.Write "<td bgcolor=""#C0C0C0"">&nbsp;</td>"
            next
         end if
         %>
      </tr>
      <tr>
      <td colspan=7>
<b>
<font color="#FF0000">���G����C</font>
<font color="#000000">���G�ư��C</font>
<font color="#0000FF">��G�f���C</font>
<font color="#BB5500">���G�����C</font>
<font color="#000000">���G�ల�C</font>
<font color="#00FF00">���G�~���C</font>
<font color="#FF0088">���G�S��C</font>
<font color="#EE7700">���G�����C</font>
<font color="#BBBB00">���G�B���C</font><BR>
<font color="#000000">���G�����d�C</font>
<font color="#FF0000">���G�ƯZ�C</font>
<font color="#000000">()�G�а��ɼơC</font>
</b>
</td>
      </tr>
   </table>

</form>
<%

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>    
</center>
</body>
</html>
