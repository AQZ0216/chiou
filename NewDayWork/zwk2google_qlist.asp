<%@ Language=VBScript CODEPAGE=950 %>

<%
'Ū���d�߱����� 
querystr=" where "
querystra=""
querystrb="�d�߱���G"

'�D��p_wk_item
p_wk_item=request("p_wk_item")
if p_wk_item="" or p_wk_item="����" then
	p_wk_item="����"
else
	querystra=querystra & "wk_item like '%"& p_wk_item &"%' and "
	querystrb=querystrb & "[�D��="& trim(p_wk_item) &"]"
	querystrc=querystrc & "p_wk_item"& trim(p_wk_item) &"&"
end if

'����H��p_wk_exe
p_wk_exe=trim(request("p_wk_exe"))
'p_wk_exe="���F"	
if p_wk_exe="" or p_wk_exe="����" then
	p_wk_exe="����"
else
	querystra=querystra & "(wk_exe like '%"& p_wk_exe &"%') and "
'	querystra=querystra & "(wk_exe like '%"& p_wk_exe &"%' or wk_exe like '����H��' ) and "
	querystrb=querystrb & "[����H��="&trim(p_wk_exe)&"]"
	querystrc=querystrc & "p_wk_exe="&trim(p_wk_exe)&"&"
end if

'���|�H��p_wk_doer
p_wk_doer=trim(request("p_wk_doer"))
'p_wk_doer="���F"	
if p_wk_doer="" or p_wk_doer="����" then
	p_wk_doer="����"
else
	querystra=querystra & "(wk_doer like '%"& p_wk_doer &"%') and "
	querystrb=querystrb & "[���|�H��="&trim(p_wk_doer)&"]"
	querystrc=querystrc & "p_wk_doer="&trim(p_wk_doer)&"&"
end if

'������p_doing_date1a
p_doing_date1a=trim(request("p_doing_date1a"))	
'p_doing_date1a="2016/3/1"
if p_doing_date1a="" or p_doing_date1a="����" then
	p_doing_date1a="����"
else
	querystra=querystra & "(doing_date1 >= #"& p_doing_date1a &"# ) and "
	querystrb=querystrb & "[������="&trim(p_doing_date1a)&"]"
	querystrc=querystrc & "p_doing_date1a="&trim(p_doing_date1a)&"&"
end if

'������p_doing_date1b
p_doing_date1b=trim(request("p_doing_date1b"))	
'p_doing_date1b="2016/4/1"
if p_doing_date1b="" or p_doing_date1b="����" then
	p_doing_date1b="����"
else
	querystra=querystra & "(doing_date1 <= #"& p_doing_date1b &"# ) and "
	querystrb=querystrb & "[������="&trim(p_doing_date1b)&"]"
	querystrc=querystrc & "p_doing_date1b="&trim(p_doing_date1b)&"&"
end if

	querystr=querystr & querystra
	len_a=len(querystr)
	if len_a=7 then querystr=" "
      if trim(querystr)="where" then querystr=" "
	if right(querystr,4)="and " then querystr=left(querystr,len_a-4)
	len_c=len(querystrc)
	if right(querystrc,1)="&" then querystrc=left(querystrc,len_c-1)
	
if trim(querystrc)="" or isnull(trim(querystrc)) then querystrc="p_wk_item=����"
	qstrURL="zwk2google_qlist.asp?"&querystrc
'�]�wsession backURL
strbackURLcsv="zwk2google_qlist_csv.asp?"&querystrc
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
<%
'response.write "strbackURLcsv="&strbackURLcsv
'response.write "<hr>"
'response.write "<a href='"& strbackURLcsv &"' target='_blank'>google csv</a>"
'response.write "<hr>"
%>
<form name="form1" method="post" action="zwk2google_qlist_csva.asp" >
	<input type="button" name="sentb" class="cbutton" value="�T�w�ץX" onclick="Verify_chk()" >
	<input type="reset" name="reset" class="cbutton" value="�M�����"  >
	<input type=button name=giveup class="cbutton" value="�^�W�@��" onclick="history.back()"  >
<hr>
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
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'strSQL_show="Select * from " & tb_name & " where wk_class like '%"&wk_class&"%' and wk_doer like '%"& wk_man &"%' order by doing_date1 desc"
strSQL_show="Select * from " & tb_name & querystr &" order by doing_date1 asc"
'	response.write strSQL_show &"<br>"
rstObj1.open strSQL_show,conDB,1,1
totalput=rstObj1.recordcount
if totalput=0 then
	str_00="Subject"'���ʦW�� (���n)�C
	str_01="Start Date"'���ʪ��Ĥ@�� (���n)�C
	str_02="Start Time"'���ʶ}�l�ɶ��C
	str_03="End Date"'���ʪ��̫�@�ѡC
	str_04="End Time"'���ʵ����ɶ��C
	str_05="All Day Event"'�o�Ӭ��ʬO�_�����Ѭ��ʡC�p�G�O���Ѭ��ʡA�п�J True�F�_�h�п�J False�C
	str_06="Description"'���ʻ����Ϊ����C
	str_07="Location"'���ʦa�I�C
	str_08="Private"'�o�Ӭ��ʬO�_���p�H���ʡC�p�G�O�p�H���ʡA�п�J True�F�_�h�п�J False�C
'	Response.Write str_00 & "," & str_01 & "," & str_02 & "," & str_03 & "," & str_04 & "," & str_05 & "," & str_06 & "," & str_07 & "," & str_08 & "<br>"
	Response.Write "�L��ƥi�ץX�C"
else
	str_00="Subject"'���ʦW�� (���n)�C
	str_01="Start Date"'���ʪ��Ĥ@�� (���n)�C
	str_02="Start Time"'���ʶ}�l�ɶ��C
	str_03="End Date"'���ʪ��̫�@�ѡC
	str_04="End Time"'���ʵ����ɶ��C
	str_05="All Day Event"'�o�Ӭ��ʬO�_�����Ѭ��ʡC�p�G�O���Ѭ��ʡA�п�J True�F�_�h�п�J False�C
	str_06="Description"'���ʻ����Ϊ����C
	str_07="Location"'���ʦa�I�C
	str_08="Private"'�o�Ӭ��ʬO�_���p�H���ʡC�p�G�O�p�H���ʡA�п�J True�F�_�h�п�J False�C
'	Response.Write str_00 & "," & str_01 & "," & str_02 & "," & str_03 & "," & str_04 & "," & str_05 & "," & str_06 & "," & str_07 & "," & str_08 & "<br>"
%>
<table border=1 style="width:1000px;">
	<col style="width:30px;text-align:center;">
	<col style="width:160px;text-align:center;">
	<col style="width:80px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:80px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:;text-align:center;">
	<col style="width:100px;text-align:center;">
	<col style="width:60px;text-align:center;">
<tr>
	<td align=center colspan=10>�@��<%=totalput%>����ƥi�ץX�C
	</td>
</tr>	
<tr>
	<td align=center width=30>�Ǹ�
		<input type="checkbox" name="psel_wkid" value="" onclick="sel_check()" title="����Υ�����"></td>
	<td align=center >Subject</td>
	<td align=center >Start Date</td>
	<td align=center >Start Time</td>
	<td align=center >End Date</td>
	<td align=center >End Time</td>
	<td align=center >All Day Event</td>
	<td align=center >Description</td>
	<td align=center >Location</td>
	<td align=center >Private</td>
</tr>	
</table>
<div style="text-align:left;width:1020px;height:305px;overflow:auto;">
<table border=1 width=1000>
	<col style="width:30px;text-align:center;">
	<col style="width:160px;text-align:left;">
	<col style="width:80px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:80px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:60px;text-align:center;">
	<col style="width:;text-align:left;">
	<col style="width:100px;text-align:center;">
	<col style="width:60px;text-align:center;">
<%
	'�C�X��ƶ���
	rstobj1.MoveFirst
	for i=1 to totalput
	'Ū�����
		wkid=rstObj1.fields("wk_id")
		doing_date1=rstObj1.fields("doing_date1")'���u���
		wk_item=replace(trim(rstObj1.fields("wk_item")),",","�A")'�D��
		wk_item=replace(wk_item,";",":")'�D��
		wk_content=left(rstObj1.fields("wk_content"),200)'�u�@���e���O
		wk_content=replace(wk_content,",","�A")'�u�@���e���O
		wk_content=replace(wk_content,chr(13),"�C")'�u�@���e���O
		wk_content=replace(wk_content,chr(10),"")'�u�@���e���O
		str1_02a=left(wk_item,5)
		if not(isnumeric(left(str1_02a,2))) then
			str1_02a="08:00"
		end if
		str1_04a=Mid(wk_item,7,5)
		if not(isnumeric(left(str1_04a,2))) then
			str1_04a=str1_02a
		end if
	str1_00=wk_item	'���ʦW�� (���n)�C
	str1_01=doing_date1		'���ʪ��Ĥ@�� (���n)�C
	str1_02=str1_02a	'���ʶ}�l�ɶ��C
	str1_03=doing_date1		'���ʪ��̫�@�ѡC
	str1_04=str1_04a		'���ʵ����ɶ��C
	str1_05="False"				'�o�Ӭ��ʬO�_�����Ѭ��ʡC�p�G�O���Ѭ��ʡA�п�J True�F�_�h�п�J False�C
	str1_06=wk_content	'���ʻ����Ϊ����C
	str1_07="taipei"'���ʦa�I�C
	str1_08="False"'�o�Ӭ��ʬO�_���p�H���ʡC�p�G�O�p�H���ʡA�п�J True�F�_�h�п�J False�C
'	Response.Write str1_00 & "," & str1_01 & "," & str1_02 & "," & str1_03 & "," & str1_04 & "," & str1_05 & "," & str1_06 & "," & str1_07 & "," & str1_08 & "<br>"
%>
<tr>
	<td><!--�Ǹ�--><%=i%>
			<input type="checkbox" name="p_wkid" value="<%=wkid%>" ><%'=wkid%>
		</td>
	<td><!--Subject--><%=str1_00%></td>
	<td><!--Start Date--><%=str1_01%></td>
	<td><!--Start Time--><%=str1_02%></td>
	<td><!--End Date--><%=str1_03%></td>
	<td><!--End Time--><%=str1_04%></td>
	<td><!--All Day Event--><%=str1_05%></td>
	<td><!--Description--><%=left(str1_06,50)%></td>
	<td><!--Location--><%=str1_07%></td>
	<td><!--Private--><%=str1_08%></td>
</tr>	
<%
	'����U�@���O��
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>
</table>
</div>
<%
end if
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
<hr>
	<input type="button" name="sentb" class="cbutton" value="�T�w�ץX" onclick="Verify_chk()" >
	<input type="reset" name="reset" class="cbutton" value="�M�����"  >
	<input type=button name=giveup class="cbutton" value="�^�W�@��" onclick="history.back()"  >
<hr>
</form>
<script language=vbscript>
sub Verify_chk()
	set celem=document.form1.getElementsByTagName("input")
	checksel1=0
	str_err=""
 	for i=0 to celem.length-1
		if celem(i).type="checkbox" and celem(i).name="p_wkid" and celem(i).checked=true then
			checksel1=checksel1+1
		end if  
	next
	if checksel1=0 then str_err=str_err&chr(13)&"�п�ܭn�ץX����ơI�I" 
	if str_err="" then
		strmsg="�T�w�ץX��ơH"& chr(13) &" �ثe�w���"& checksel1 &"����ơI�I"
		chkq=msgbox(strmsg,64+1,"�T�{�T��")
		if chkq=1 then
			document.form1.submit
		else
		end if
	else 	 
		errcode=msgbox("���~�T���I�I�I"& chr(13)&str_err& chr(13) ,64+0,"���~�T��")
	end if
end sub
sub sel_check()
	set celema=document.form1.getElementsByTagName("input")
	if document.form1.psel_wkid.checked=true then
		check_all
	else
		uncheck_all
	end if  
end sub

sub check_all()'����
	set celem=document.form1.getElementsByTagName("input")
 	for i=0 to celem.length-1
		if celem(i).type="checkbox" and celem(i).name="p_wkid" then
			celem(i).checked=true
		end if  
	next
end sub
sub uncheck_all()'������
	set celem=document.form1.getElementsByTagName("input")
 	for i=0 to celem.length-1
		if celem(i).type="checkbox" and celem(i).name="p_wkid" then
			celem(i).checked=false
		end if  
	next
end sub
</script>
</body>
</html>