<%@ Language=VBScript CODEPAGE=950 %>
<%

'�C����ܵ��Ƭ�100�� data_no
data_no=100
'�ثe���X page_no
if request("page_no")="" then
	page_no=1 
else
	page_no=request("page_no")
end if

%>
<!-- �}�Ҹ�Ʈw -->
<%
' �s��Access��Ʈw./database/daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="project_data"

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " order by pj_01 desc"
rstObj1.open strSQL_show,conDB,3,3
'�p�����`��	
totalput=rstObj1.recordcount	
%>	 

<html>
<head>
<title>�M�צC��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--Ū�J�ù���ܼ˪O�� base_screen_�@��.css �ΦC�L�˪O�� base_print_�@��.css  -->
	<link rel="stylesheet" type="text/css" 
		media="screen" href="./css/base_screen.css" title="style_screen">
<!--�]�w�˪O�榡-->
<style type="text/css">
	<!--

	-->
</style>
</head>
<body>

<!-- ���D -->
<%
if totalput=0 then
%>
<font style="font-family:�s�ө���;font-size:5mm;font-weight:bold;color:#006400;">
�L�M�סI�I
</font> 
<%
else
	data_ck=totalput mod data_no
	if data_ck=0 then
		page_total=int(totalput/data_no)
	else
		page_total=int(totalput/data_no)+1
	end if
%>
<%
if page_total=1 then
		page_no_b=1
		page_no_g=1
else
	if page_no=1 then
		page_no_b=1
		page_no_g=page_no+1
	else
		if cint(page_no)=cint(page_total) then
			page_no_b=page_no-1
			page_no_g=page_total
		else
			page_no_b=page_no-1
			page_no_g=page_no+1
		end if
	end if
end if

'�p��_�l����
datafirst=(page_no-1)*data_no+1
if cint(page_no)=cint(page_total) then
	datalast=totalput
else
	datalast=datafirst+data_no-1
end if
'�p�⥻������ 
no_local=datalast-datafirst+1
'�]�wsession backURL
strbackURLa="pj_list.asp?page_no="
strbackURL=strbackURLa&page_no
Session("strbackURL")=strbackURL

%>
<font style="font-family:�s�ө���;font-size:3.0mm;font-weight:normal;">
<table border=0 style="width:750px;">
<tr style="height:35px;">
<td style="width:20%;text-align:center;background-color:#e0e0e0;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#e0e0e0';">
<a href="pj_add.asp" target="_self" style="text-decoration:none;color:blue;">�i�s�W�M�סj</a>
</td> 
<td style="width:60%;text-align:center;">
<font style="font-family:�s�ө���;font-size:4mm;font-weight:bold;color:#006400;">
�Ҧ��m�M�׽s���n�C��
</font>
<%
for j=1 to page_total
%>
<a href="<%=strbackURLa%><%=j%>" ><%=j%>&nbsp;</a>
<%
next
%>
<td style="width:20%;text-align:center;background-color:#e0e0e0;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#e0e0e0';">
<a href="pj_list_print.asp?page_no=<%=page_no%>" target="_self" style="text-decoration:none;color:blue;">�i�C�L�����j</a>
</td> 
</table>
</font>

<div style="text-align:left;width:775px;height:55px;overflow:off;">
<!-- ��ƦC����D�}�l -->
<font style="font-family:�s�ө���;font-size:3.5mm;font-weight:normal;">
<table border=0 style="font-size:3.5mm;text-align:left;width:750px;">
<tr>
<a href="<%=strbackURLa%><%=page_no_b%>">
	<td style="width:8%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		<a href="<%=strbackURLa%><%=page_no_b%>">�e�@��<br>(��<%=page_no_b%>��)</a>
	</td>
</a>
	<td style="font-size:4mm;width:8%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		��<font color=red><%=page_no%></font>��
	</td>
<a href="<%=strbackURLa%><%=page_no_g%>">
	<td style="width:8%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		<a href="<%=strbackURLa%><%=page_no_g%>">�U�@��<br>(��<%=page_no_g%>��)</a>
	</td>
</a>
	<td style="text-align:center;">
		<font style="font-family:�s�ө���;font-size:3.5mm;font-weight:normal;">
		�ثe���X�O��<font color=red>&nbsp;<%=page_no%>&nbsp;</font>���A�@��<font color=red>&nbsp;<%=page_total%>&nbsp;</font>��(�C��<%=data_no%>��)�A�@��<font color=red><%=totalput%></font>�����<br> 
		������Ƭ���<font color=red>&nbsp;<%=datafirst%>&nbsp;��&nbsp;<%=datalast%>&nbsp;</font>���A�@<font color=red><%=no_local%></font>�����
		</font>
<a href="<%=strbackURLa%><%=page_no_b%>">
	<td style="width:10%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		<a href="<%=strbackURLa%><%=page_no_b%>">�e�@��<br>(��<%=page_no_b%>��)</a>
	</td>
</a>
	<td style="font-size:4mm;width:8%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		��<font color=red><%=page_no%></font>��
	</td>
<a href="<%=strbackURLa%><%=page_no_g%>">
	<td style="width:10%;text-align:center;background-color:#ffdab9;cursor:hand;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#ffdab9';">
		<a href="<%=strbackURLa%><%=page_no_g%>">�U�@��<br>(��<%=page_no_g%>��)</a>
	</td>
</a>
</table>
</font>
<font style="font-family:�s�ө���;font-size:4mm;font-weight:normal;">
<!-- ��ƦC��e�����e -->
<input type='hidden' name="firstdata" value="<%=datafirst%>">
<input type='hidden' name="lastdata" value="<%=datalast%>">
<input type='hidden' name="local_no" value="<%=no_local%>">
<!-- ��ƦC���e�}�l -->
<font style="font-family:�s�ө���;font-size:3.5mm;font-weight:normal;">
	<table border=0 style="text-align:left;width:750px;">
	<col style="font-size:3.5mm;width:60px;text-align:center;">
	<col style="font-size:3.5mm;width:80px;text-align:center;">
	<col style="font-size:3.5mm;width:120px;text-align:center;">
	<col style="font-size:3.5mm;;text-align:center;">
	<tr>
		<td>�\����</td> 
		<td>�M�׽s��
		<td>�M�צW�� 
		<td>�M�׻��� 
	</tr>
	</table>
</font>
</div>
<div style="text-align:left;width:775px;height:255px;overflow:auto;">
<font style="font-family:�s�ө���;font-size:3.5mm;font-weight:normal;">
	<table border=0 style="text-align:left;width:750px;">
	<col style="font-size:3.5mm;width:60px;text-align:center;">
	<col style="font-size:3.5mm;width:80px;text-align:center;">
	<col style="font-size:3.5mm;width:120px;text-align:center;">
	<col style="font-size:3.5mm;;text-align:left;">

<%

'���ܲĤ@����� 
rstobj1.MoveFirst
'���ܰ_�l���� 
rstobj1.move datafirst-1

%>
	<%	
	'�C�X��ƶ���
	'rstobj1.MoveFirst
	for j=datafirst to datalast
	'�]�w�ťո�Ƥ��ϬM
p_id=rstObj1.fields("pj_id")	'�M��id
p_01=rstObj1.fields("pj_01")	'�M�׽s��
p_02=rstObj1.fields("pj_02")	'�M�צW��
p_03=rstObj1.fields("pj_03")	'�M�׻���

oddchk = j mod 2
if oddchk=1 then
	BKC="#ddffee"
else
	BKC="#ffffee"
end if
	%>
	<tr style="background-color:<%=BKC%>;" onmouseover="javascript:this.style.background='#FFeedd';" onmouseout="javascript:this.style.background='<%=BKC%>';">
	<td valign=middle >
	<a href="./pj_del.asp?p_id=<%=p_id%>"> <img src="./img/del1.gif" alt="�R���M��" width="15" height="15" style='cursor:hand;border:0;'></a>
	<a href="./pj_edit.asp?p_id=<%=p_id%>"> <img src="./img/edit1.gif" alt="�s��M��" width="15" height="15" style='cursor:hand;border:0;'></a>
	</td>
	<td><%=p_01%></td>
	<td><a href="./pj_show.asp?p_id=<%=p_id%>" ><%=p_02%></a></td>
	<td>&nbsp;&nbsp;<%=p_03%></td>
</tr>	
	
	<%
	'����U�@���O��
				
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next
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
<!-- ��ƦC����-->	
</table>
</font>
</div>

</body>
</html> 

