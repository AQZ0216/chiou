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
   <div class="w3-row w3-pale-green ">
      <!--�\���1-->
      <button class="w3-button w3-red w3-medium w3-round" onclick="url_new('http://192.168.0.11/chiou/att2000/5_card_query.asp')" title="�X�Ԯɶ�" style="padding:4px;margin:2px;">�X�Ԯɶ�</button>
      <button class="w3-button w3-red w3-medium w3-round" onclick="url_new('../holiday/hd_ps_year_list.asp?wkr_id=<%=pwkr_id%>')" title="�𰲸��" style="padding:4px;margin:2px;">�𰲸��</button>
      <button class="w3-button w3-red w3-medium w3-round" onclick="url_new('http://1.34.48.166:90/firstpage.asp')" title="�y�����" style="padding:4px;margin:2px;">�y�����</button>
      <button class="w3-button w3-red w3-medium w3-round" onclick="url_new('http://60.251.159.62:6980/build/daywork/firstpage.asp?paswd=28283939')" title="�س]��" style="padding:4px;margin:2px;">�س]��</button>
   </div>  
</div>
<!--���e-->
<div class="w3-brown w3-center" >
   <div class="w3-row w3-pale-blue ">
      <!--�\���2-->
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-wk_Calendar_all.asp?worker=<%=worker%>')" title="�^����" style="padding:4px;margin:2px;">�^����</button>
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-work_query.asp?worker=<%=worker%>')" title="�u�@�d��" style="padding:4px;margin:2px;">�u�@�d��</button>
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-work_add.asp?worker=<%=worker%>')" title="�u�@�s�W" style="padding:4px;margin:2px;">�u�@�s�W</button>
   </div>     
   <div class="w3-row w3-center w3-pale-blue" style="width:100%;height:550px;overflow: hidden;">
      <iframe id="ifrm_content" name="ifrm_content" src="zec-wk_Calendar_new.asp?worker=<%=worker%>" style="border:0px;width:100%;height:100%;"></iframe>	
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