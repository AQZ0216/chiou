<% @codepage=950%>
<%
	'Ū���H���m�W
	worker = Session("worker")
	pwd_headline=Request("pwd_headline")   '�K�X
%>

<%
if pwd_headline="3939" then
      '�N�u�@�C�����j�T��
      ' �s��Access��Ʈwdaywork.mdb
      DBpath=Server.MapPath("./database/daywork.mdb")
      strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
      '�إ߸�Ʈw�s������
      set conDB= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB.Open strCon
      '�}�Ҹ�ƪ�W��
      tb_name="work_data"
      '�إ߸�Ʈw�s������	
      set rstObj1=Server.CreateObject("ADODB.Recordset")
      strSQL_show="Select * from " & tb_name & " where headline > 5  order by doing_date1 asc"
      rstObj1.open strSQL_show,conDB,1,3
totalput=rstObj1.recordcount
if totalput=0 then
else
	'�C�X��ƶ���
	rstobj1.MoveFirst
	for i=1 to totalput
         rstObj1.fields("headline")=5
	'����U�@���O��
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
end if
      rstObj1.UpdateBatch
      '������ƶ�
      rstObj1.Close
      '���]����ܼ�
      set rstObj1=Nothing
      '������Ʈw
      conDB.Close
      '���]�����ܼ�
      set conDB=Nothing
    strURL1=session("hback_URL")
      'strURL1="wk_lst_doing.asp"
      response.redirect(strURL1)

else
    if isnull(pwd_headine) or pwd_headline="" then
         pwd_msg=""
    else
         pwd_msg="�K�X���~�I�I"
    end if
%>

<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<form name="form1" action="0_wk_headline_off_all.asp" method="post">
<input type="hidden" name="worker" value="<%=worker%>">
<font style="font-size:16pt;" color="red">�N�Ҧ����j�T�������A�п�J�K�X�I�I</font><br>
<font style="font-size:12pt;" color="blue">�N�����]���O���Ҧ����j�T�������I�I</font><br>
�K�X�G<input type='password' name='pwd_headline' value='' style="width:100px;" > <br>
		<input type=submit name="editok" value="�T�w" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100px;">
		<input type=button name="goback1" value="�^�W�@��" onclick="history.back()" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100px;" >
<hr>
<%
if pwd_msg="�K�X���~�I�I" then
      response.write "<b>" & pwd_msg & "</b><hr>"
end if
%>

</form>
<center>
</body>
</html>
<% end if %>