<% @codepage=950%>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<%
   '�]�wSession�ܼƮ����ɶ�
   Session.TimeOut=480
'Session("worker")=Request("worker")
worker_old = Session("worker")
if request("fp")="1" then worker_old="��j"
'if worker_old="" or isnull(worker_old) then worker_old="��j"
worker=Request("worker")
Session("worker")=worker
%>
<%
'Ū���K�X���
Function find_crewpwd(p_wkr)
      ' �s��Access��Ʈw../holiday/database/crew.mdb
      DBpath_p=Server.MapPath("../holiday/database/crew.mdb")
      strCon_p="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_p
      '�إ߸�Ʈw�s������
      set conDB_p= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB_p.Open strCon_p
      '�}�Ҹ�ƪ�W��
      tb_name_p="crew"
      '�إ߸�Ʈw�s������	
      set rstObj1_p=Server.CreateObject("ADODB.Recordset")
      strSQL_show_p="Select * from " & tb_name_p & " where worker like '" & p_wkr &"'"
      rstObj1_p.open strSQL_show_p,conDB_p,3,1
      totalp=rstObj1_p.recordcount
      if totalp=0 then
         p_pwd="0"
      else
         p_pwd=rstObj1_p.fields("wkr_pwd")		'�K�X
      end if
      '������ƶ�
      rstObj1_p.Close
      '���]����ܼ� 
      set rstObj1_p=Nothing
      '������Ʈw
      conDB_p.Close
      '���]�����ܼ� 
      set conDB_p=Nothing
      find_crewpwd=p_pwd
End Function
%>
<%
'---------------------------------------------------------------------------
wkr_pwd=session("wkr_pwd") 'Ū���K�X
if session("wkr_pwd")="" or isnull(wkr_pwd) then
      wkr_pwd=request("wkr_pwd") 'Ū���K�X
end if
'chk_str="���`�B���B����B���z�B�f�S"
chk_str=""
'if instr(1,chk_str,worker,1)>0 and instr(1,chk_str,worker_old,1)=0 then
if instr(1,chk_str,worker,1)>0 then
   if instr(1,chk_str,worker_old,1)>0 then
   'response.write "worker_old="&worker_old
   'response.end
   else
         'Ū����Ʈw�K�X
      '   response.write "worker="&worker&"<br>"
         db_pwd=find_crewpwd(worker)
         if db_pwd=wkr_pwd then
           session("wkr_pwd")=db_pwd
         else
            str_url="./0_login_pwd.asp?worker="&worker
            response.redirect(str_url)      '��}��K�X��J�e��
         end if
      end if
end if
'---------------------------------------------------------------------------
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
		       pstr_hdman = pstr_hdman & "<span style='font-size:14px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man & "</span><br>"
   	           end if
		else
		  pstr_hdman = pstr_hdman & "<span style='font-size:14px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man &"("& hd_hrs&")&nbsp;</span><br>"
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
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<title>�i<%=worker%>�j�u�@�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="./css/w3-cht.css">
<style>

</style>
</head>
<body class="vt-container">
<!-- ���Y�}�l -->
<!-- #Include file = "./include/zec-header_r1.inc" -->
<!-- ���Y���� -->
<!-- ����}�l -->
<div class="vt-container w3-pale-red w3-center" >
   <div class="w3-row w3-center " ><!-- ���e start -->
<%
datecode=request("datecode")
p_year=year(datecode)'�~
p_month=month(datecode)'��
p_day=day(datecode)'��
pn_date=dateserial(p_year,p_month,p_day)'�d�ߤ��
pn_weekday=Weekday(pn_date)'�P���X

      Select Case pn_weekday
         Case 1    
            str_wk="�P����"
            div_background_c="#ffdddd"    'w3-pale-red
            div_border_c="#000"        'w3-pale-red
         Case 2    
            str_wk="�P���@"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
         Case 3    
            str_wk="�P���G"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
        Case 4    
            str_wk="�P���T"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
         Case 5    
            str_wk="�P���|"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
         Case 6    
            str_wk="�P����"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
        Case 7    
            str_wk="�P����"
            div_background_c="#ffdddd"    'w3-pale-red
            div_border_c="#000"        'w3-pale-red
         Case Else     
      End Select 

%>
      <table class="w3-table-all" >
         <col style="width:100%;background-color:#ffdddd;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <td style="text-align:center;height:55px;">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�W�@�~" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=cint(p_year)-1%>&p_month=<%=p_month%>')" >�i<<�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�W�@��" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=p_year%>&p_month=<%=cint(p_month)-1%>')" >�i<�j</button>
            <button class="w3-button w3-white w3-xlarge " style="padding:2px;margin:0px;" >�i<%=p_year%>�~<%=p_month%>��<%=p_day%>��j���u�@����</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�U�@��" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=p_year%>&p_month=<%=cint(p_month)+1%>')">�i>�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�U�@�~" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=cint(p_year)+1%>&p_month=<%=p_month%>')">�i>>�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�^<%=year(date())%>�~<%=month(date())%>��" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=year(date())%>&p_month=<%=month(date())%>')">�i����G<%=year(date())%>�~<%=month(date())%>��j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�^<%=year(date())%>�~<%=month(date())%>��<%=day(date())%>��" onclick="url_show('zec-work_day_r1.asp?worker=<%=worker%>&datecode=<%=date()%>')">�i����G<%=date()%>�j</button>
           </td>
         </tr>
      </table>
<div class="w3-container" style="margin:0px;padding:0px;height:520px;overflow:auto;"><!-- ���� start -->
      <div class="w3-responsive">
      <table class="w3-table-all" >
         <col style="width:100%;background-color:<%=div_background_c%>;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <th style="text-align:center;background-color:<%=div_background_c%>;border:1px solid #000;">
           <span class="w3-button" style="font-size:24px;padding:0px;margin:0px;">�i<%=p_year%>�~<%=p_month%>��<%=p_day%>��j<%=str_wk%></span>
           <span class="w3-button w3-red" title="�s�W�u�@" style="font-size:24px;padding:0px;margin:0px;" onclick="url_show('zec-work_add_r1.asp?datecode=<%=pn_date%>&worker=<%=worker%>')" > �i�s�W�j </span>
           </th>
         </tr>
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
<%
      pndate=pn_date '��ܪ��u�@���
      str_hdman=hd_man(pndate)'�𰲦r��
'-------------------------str_allwork----------------------------
      str_allwork="" '�u�@�r��
      '�إ߸�Ʈw�s������	
      set rstObj1=Server.CreateObject("ADODB.Recordset")
      strSQL_show="Select * from " & tb_name & " where doing_date1 = #"& pndate &"# and wk_undoer like '%"& worker &"%' order by wk_item asc , wk_id asc"
      rstObj1.open strSQL_show,conDB,3,1
      tot_d1=rstObj1.recordcount
      if not rstObj1.EOF then
         rstObj1.Movefirst
         for di = 1 to tot_d1
            p_wkid=rstObj1.fields("wk_id")
            p_wkitem=rstObj1.fields("wk_item") 
            '----------------------------------------
               wk_headline=rstObj1.fields("headline")  '�]���O
               '�ˬd�O�_������ exist_attach(wk_id)
               attach_no=exist_attach(p_wkid)
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
            '----------------------------------------
           
            str_allwork = str_allwork & "<span style='font-size:14px;"& str_bgc & str_colors &"' >" & di &"�B<a href='zec-work_show_r1.asp?wk_id="& p_wkid &"&worker="& worker&"' style='text-decoration: none;"& str_colors &"' >" & p_wkitem &"</a></span><br>" 
            rstObj1.MoveNext
            if rstObj1.EOF=true then exit for
         next
      else
      end if
      '������ƶ�
      rstObj1.Close
      '���]����ܼ� 
      set rstObj1=Nothing
'-------------------------str_allwork----------------------------
%>          
   <tr>
      <td style="background-color:<%=div_background_c%>;border-color:<%=div_border_c%>;" >
         <%=str_allwork%><%=str_hdman%>       
      </td>
   </tr>
   </table>
      </div>
<%

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%> 	      
</div><!-- ���� end -->

   </div><!-- ���e end -->
</div>
<!-- ���嵲�� -->
<!-- �����}�l -->
<!-- #Include file = "./include/zec-footer_r1.inc" -->
<!-- �������� -->

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
    function content_show(pp_url){
        var iframe1=document.getElementById("ifrm_content");
        iframe1.src=pp_url;
        return true;
    }    

</script>

</body>
</html>