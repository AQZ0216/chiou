<%@ Language=VBScript CODEPAGE=950 %>

<HTML>
<HEAD>
<title>�ק�����s�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" 
href="base_first.css" title="style1">
<style type="text/css"><!--
.ma1{
	font-family:'�s�ө���';
	color:red;
	font-size:12pt;
	} 
.ma2{
	font-family:'�s�ө���';
	color:black;
	font-size:10pt;
	} 
.ma3{
	font-family:'�s�ө���';
	color:black;
	font-size:10pt;
	} 
--></style>
<script language="JavaScript">
</script>        
</HEAD>
<BODY topmargin=5>
<CENTER>
<font class='tit'>�ק�����s�����</font>
<%
' �s��Access��Ʈw./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="linkdata"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " order by lk_row asc, lk_col asc"
rstObj1.open strSQL_show,conDB,3,1
'�p�����`��	
totalput=rstObj1.recordcount
if totalput=0 then
else
%>
<table border="1" cellspacing=0 cellpadding=0 width=783>
<col span=6 style="width:16.6%;">
<%
      '���ܲĤ@�����
      rstobj1.MoveFirst
      p_04old=0
      '�C�X��ƶ���
      for i=1 to totalput
      	'�]�w�ťո�Ƥ��ϬM
      p_id=rstObj1.fields("lk_id")		'id	
      p_01=rstObj1.fields("lk_url")		'�s�����}
      p_02=rstObj1.fields("lk_item")		'�u���D
      p_03=rstObj1.fields("lk_title")		'�y�z
      p_04=rstObj1.fields("lk_row")		'�C
      p_05=rstObj1.fields("lk_col")		'��
if p_02="" or isnull(p_02) then
   p_02="--"
   p_03="�ק�s�����}"
end if

if p_04=p_04old then
else
      if p_04<>1 then response.Write   "</tr>"
      response.Write   "<tr align=center style='height:20pt;' >"
      p_04old=p_04
end if

   if len(p_02)>7 then
      str_ft="font-size:11pt;"
   else
      str_ft="font-size:12.5pt;"
   end if
%>
<A Href='0_edit_link.asp?row=<%=p_04%>&col=<%=p_05%>' target='_self' ><td class=urlcmd title='<%=p_03%>' style='<%=str_ft%>'><%=p_02%></td></A>
<%

%>

<%


      '����U�@���O��
      rstObj1.MoveNext
      if rstObj1.EOF=True then exit for
      next

   response.Write   "</tr>"
end if

'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing
new_row=p_04+1
%>
</table>
<hr>
<a href="0_new_row.asp?new_row=<%=new_row%>">�s�W�@�C</a> &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp;
<a href="firstpage.asp">�^����</a>&nbsp;  &nbsp; &nbsp; &nbsp; &nbsp;
<a href="0_del_row.asp?last_row=<%=p_04%>">�R���̫�@�C</a>
</center>

</BODY>
</HTML>






