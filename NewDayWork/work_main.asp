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
<html>

<head>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="icon" href="../daywork/img/khouse.ico" type="image/ico" />
</head>

<frameset rows="85,*" >
  <frame name="topfrm" src="work_tit.asp?worker=<%=worker%>" marginwidth=0 marginheight=0 scrolling="no" >
<!--  <frame name="content" src="wk_lst_doing.asp?worker=<%=worker%>" scrolling="auto">-->
<frame name="content" src="wk_Calendar_all.asp?worker=worker" scrolling="auto">
</frameset>
</html>
