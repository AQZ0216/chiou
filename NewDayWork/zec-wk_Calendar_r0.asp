<%@ Language=VBScript CODEPAGE=950 %>
<%
   'Ū���H���m�W
   worker = request("worker")
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
		   Case 17  f_color = "#000000"    '�I�G���˰��C
		   Case 18  f_color = "#000000"    '�I�G�������C
		   Case 19  f_color = "#000000"    '��G�|�����C
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
nMonth = Request("nMonth")
nYear = Request("nYear")
if nMonth = "" then nMonth = Month(dtDate)
if nYear = "" then nYear = Year(dtDate)
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
dtDate = DateSerial(nYear, nMonth, 1)


if int(nMonth)<10 then
   strnMonth="0"&cstr(nMonth)
else
   strnMonth=cstr(nMonth)
end if
dcodeym=cstr(nYear)&strnMonth

'�]�wsession("strbackURL")
strbackURL="wk_Calendar_all.asp?nMonth="&nMonth&"&nYear="&nYear
session("strbackURL")=strbackURL

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

<HTML>
<HEAD>
<title>�˪O���D</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="./css/w3-cht.css">
<style type="text/css">
<!--
div.dayblock{
   border-collapse:collapse; 	/*��اΦ����X*/
	font-family:�L�n������;
	/*letter-spacing:2px;*/
	font-size:12px;
	font-weight:bold;
	color:#000000;
	/*cursor:hand;*/
	background-color:#fcfcfc;
	border: 5px solid #fcfcfc;
	/*margin:1px;*/	
   width:100%;
   /*height:150px;*/
   min-height:150px;
   /*max-height:200px;*/
   /*overflow: auto;*/
	}
-->
</style>

</HEAD>
<body class="vt-container w3-pale-blue" style="overflow:hidden;">
<center>
<form method="post" name="form1" action="">
<div class="w3-pale-blue w3-center" >
   <!--�\���-->
   <div class="w3-bar w3-blue" >
      <button onclick="url_show('zec-wk_Calendar_r0.asp?worker=<%=worker%>')" class="w3-bar-item w3-button w3-mobile" style="padding:4px;margin:0px;">�^����</button>
      <button onclick="url_show('zec-work_query.asp?worker=<%=worker%>')" class="w3-bar-item w3-button w3-mobile" style="padding:4px;margin:0px;">�u�@�d��</button>
      <button onclick="url_show('zec-work_add.asp?worker=<%=worker%>')" class="w3-bar-item w3-button w3-mobile" style="padding:4px;margin:0px;">�u�@�s�W</button>
   </div>
<!--
   <div class="w3-row w3-center w3-pale-blue ">
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-wk_Calendar_r0.asp?worker=<%=worker%>')" title="�^����" style="padding:4px;margin:2px;">�^����</button>
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-work_query.asp?worker=<%=worker%>')" title="�u�@�d��" style="padding:4px;margin:2px;">�u�@�d��</button>
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-work_add.asp?worker=<%=worker%>')" title="�u�@�s�W" style="padding:4px;margin:2px;">�u�@�s�W</button>
   </div> 
-->       
   <!--���e-->
<%
p_year=request("p_year")'�~
p_month=request("p_month")'��
p_day=request("p_day")'��
'p_week=request("p_week")'�g��
if p_year="" or isnull(p_year) then p_year=year(date())
if p_month="" or isnull(p_month) then p_month=month(date())
if p_day="" or isnull(p_day) then p_day=day(date())

pn_date=dateserial(p_year,p_month,p_day)'����

p_year=year(pn_date)'�~
p_month=month(pn_date)'��
p_week=DatePart("ww",pn_date)'�g��
p_date=pn_date'����
p_showtype="month"   'month�Bweek�Bdate
p_mfweek=DatePart("ww",dateserial(p_year,p_month,1))'����1��g��
p_mfweekday=Weekday(dateserial(p_year,p_month,1))'����Ĥ@��P��

'�w���~��+�g�ơA�d�߶g�ƲĤ@��(�P�鬰�Ĥ@��)
function findwk01(pp_yy,pp_wks)
   pp_wk01=DatePart("w",dateserial(pp_yy,1,1))'�����P���X
   pp_wk01dayno=(7-pp_wk01)+1'�Ĥ@�g�Ѽ�
   pp_dayno=(pp_wks-2)*7+pp_wk01dayno   '�����
   findwk01=DateAdd("d", pp_dayno, dateserial(pp_yy,1,1)) 
end function
p_firstday=dateserial(p_year,p_month,1-(p_mfweekday)+1)'��ܶg�ƲĤ@��
%>
      <div class="w3-row ">
         <div class="w3-col l74 w3-center w3-grey w3-border w3-border-black ">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i<<�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i<�j</button>
            <%=p_year%>�~<%=p_month%>��
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i����=<%=pn_date%>�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i>�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i>>�j</button>
         </div>
         <div class="w3-col l73 w3-center w3-grey w3-border w3-border-black ">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i��j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i�g�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i��j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i��j</button>
         </div>
      </div>
   <div class="w3-row w3-center w3-pale-red" style="max-height:460px;overflow:scroll;">
<!--
      <div class="w3-row " style="overflow:auto;">
         <div class="w3-col l71 w3-center w3-pale-red w3-border w3-border-black " style="padding:4px;margin:0px;">�P����</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">�P���@</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">�P���G</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">�P���|</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">�P���|</div>
         <div class="w3-col l71 w3-center w3-pale-green w3-border w3-border-black " style="padding:4px;margin:0px;">�P����</div>
         <div class="w3-col l71 w3-center w3-pale-red w3-border w3-border-black " style="padding:4px;margin:0px;">�P����</div>
      </div> -->
<%
for wkno=1 to 6
   pn_wkno=p_mfweek+wkno-1
   pw_day01=findwk01(p_year,pn_wkno)'���P�Ĥ@��
  
   if month(pw_day01)=p_month or month(pw_day01+6)=p_month then
      if pn_wkno=p_week then 
         div_background_c="#fffed9"
         div_border_c="#fffed9"
      else
         div_background_c="#fcfcfc"
         div_border_c="#fcfcfc"
      end if

   for dn=1 to 7
      Select Case dn
         Case 1    
            str_wk="��"
            div_background_c="#ffdddd"    'w3-pale-red
            div_border_c="#ffdddd"        'w3-pale-red
         Case 2    
            str_wk="�@"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
         Case 3    
            str_wk="�G"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
        Case 4    
            str_wk="�T"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
         Case 5    
            str_wk="�|"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
         Case 6    
            str_wk="��"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#ddffdd"        'w3-pale-green
        Case 7    
            str_wk="��"
            div_background_c="#ffdddd"    'w3-pale-red
            div_border_c="#ffdddd"        'w3-pale-red
         Case Else     
         
      End Select   

      pndate=pw_day01+dn-1'���
      
      if month(pndate)=p_month then
         '������
         'div_background_c="#fcfcfc"
         'div_border_c="#fcfcfc"
      else
         '�D������
         div_background_c="#dbdbdb"
         div_border_c="#dbdbdb"
      end if
      
      if pndate=pn_date then 
         div_background_c="#fffed9" 
         div_border_c="#fffed9"
      end if
%>   
      <div class="w3-col l71 w3-center w3-border w3-border-black dayblock" style="background-color:<%=div_background_c%>;border-color:<%=div_border_c%>;">
         <div class="w3-container" style="overflow:auto;">
            <div class="w3-col s6">
               <%=pw_day01+dn-1%>(<%=str_wk%>)
            </div>
            <div class="w3-col s6">
               <button class="w3-button w3-grey w3-medium " style="padding:0px;margin:0px;">�i�s�W�j</button>
            </div>
            <div class="w3-row w3-left " style="overflow: auto;">
               <%
               str_hdman=hd_man(pndate)
               response.write str_hdman
               %>
            </div> 
          </div>          
      </div>
<% 
   next 
   end if
next
%>
   </div>

</div>

</form>
<script language="JavaScript">
    function url_new(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        //window.location.href = pp_url; //�쭶����s
        window.open(pp_url) ; //�}�ҷs����
        return true;
    }   
    function url_show(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        window.location.href = pp_url; //�쭶����s
        //window.open(pp_url) ; //�}�ҷs����
        return true;
    }   
    function zms_show(pp_url){
        var iframe1=document.getElementById("ifrm_milestone");
        iframe1.src=pp_url;
        return true;
    }    
    function zlb_show(pp_url){
        var iframe1=document.getElementById("ifrm_logbook");
        iframe1.src=pp_url;
        return true;
    }
    function zfi_show(pp_url){
        var iframe1=document.getElementById("ifrm_finance");
        iframe1.src=pp_url;
        return true;
    }
</script>
</center>
</body>
</html>
