<% @codepage=950%>
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
<!--���Y-->
<div class="header w3-brown w3-center" style="overflow: hidden;">
  <button class="w3-button w3-brown w3-xlarge w3-round-large" onclick="location.reload()" title="���㭶��" style="padding:4px;margin:4px;">
  �i<%=worker%>�j�u�@�޲z
  </button>
</div>
<div class="w3-pale-blue w3-center" >
    <div class="w3-bar w3-green "><!-- �\���1 start -->
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('./zec-firstpage.asp')" title="�^����">�^����</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('http://192.168.0.11/chiou/att2000/5_card_query.asp')" title="�t�}�����A�X�Ԯɶ�">�X�Ԯɶ�</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('../holiday/hd_ps_year_list.asp?wkr_id=<%=pwkr_id%>')" title="�t�}�����A�𰲸��">�𰲸��</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('http://1.34.48.166:90/firstpage.asp')" title="�t�}�����A�y�����">�y�����</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('http://60.251.159.62:6980/build/daywork/firstpage.asp?paswd=28283939')" title="�t�}�����A�س]��">�س]��</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('../customer/cr_wk_con.asp?user=<%=worker%>&pswdck=1')" title="�Ȥ�d��">�Ȥ�d��</button>
   </div> <!-- �\���1 end -->
</div> 
<div class="w3-pale-red w3-center" >
   <div class="w3-bar w3-blue" ><!-- �\���2 start -->
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-wk_Calendar_r0.asp?worker=<%=worker%>')">�^����</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-work_query.asp?worker=<%=worker%>')">�u�@�d��</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-work_add.asp?worker=<%=worker%>')">�u�@�s�W</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-wk_pj_list.asp')" title="�M�פu�@">�M�פu�@</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-2_admin_main.asp')" title="��x�޲z">��x�޲z</button>
   </div><!-- �\���2 end -->
   <div class="w3-row w3-center " ><!-- ���e start -->
      <div class="w3-responsive">
      <table class="w3-table-all" border=1>
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <tr>
           <td colspan=5 style="text-align:center;">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i<<�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i<�j</button>
            <%=p_year%>�~<%=p_month%>��
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i����=<%=pn_date%>�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i>�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i>>�j</button>
           </td>
           <td colspan=2 style="text-align:center;">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i��j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i�g�j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i��j</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">�i��j</button>
           </td>
         </tr>
         <tr>
           <th style="text-align:center;">�P����</th>
           <th style="text-align:center;">�P���@</th>
           <th style="text-align:center;">�P���G</th>
           <th style="text-align:center;">�P���T</th>
           <th style="text-align:center;">�P���|</th>
           <th style="text-align:center;">�P����</th>
           <th style="text-align:center;">�P����</th>
         </tr>
         <tr>
           <td>Jill</td>
           <td>Smith</td>
           <td>50</td>
           <td>50</td>
           <td>50</td>
           <td>50</td>
           <td>50</td>
         </tr>
         <tr>
           <td>Eve</td>
           <td>Jackson</td>
           <td>94</td>
           <td>94</td>
           <td>94</td>
           <td>94</td>
           <td>94</td>
         </tr>
         <tr>
           <td>Adam</td>
           <td>Johnson</td>
           <td>67</td>
           <td>67</td>
           <td>67</td>
           <td>67</td>
           <td>67</td>
         </tr>
      </table>
      </div>
   </div><!-- ���e end -->
</div>
<!--���e-->
<div class="w3-red w3-center" >
   <div class="w3-row w3-center " >

<!--      <iframe id="ifrm_content" name="ifrm_content" src="zec-wk_Calendar_r0.asp?worker=<%=worker%>" style="border:2px;width:100%;height:100%;"></iframe>	-->
   </div>
</div>
<!--����-->
<div class="footer w3-brown w3-center"  style="height:40px;overflow: hidden;margin:0px;padding:0px;">
  <h6>�@�@�@���v�Ҧ��@�@�@<strong>��j�a���}�o�ѥ��������q</strong>�@�@�@2021�@�@�@</h6>
</div>

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