<% @codepage=950%>
<!-- #Include file = "./include/f_week_cstr.inc" -->
<%

	'Ū���H���m�W
	worker = Session("worker")
	'Ū�����Ѥ��
	ckdate=date()
wkgroup="�@��u�@"
'��ܧ����u�@���~�� 
cyear=request("cyear")
if cyear="" or isnull(cyear) then cyear=year(date()) 
dateu=cstr(cyear)&"/12/31" 
daten=cstr(cyear)&"/1/1" 
%>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_checker like '%"&worker&"%' order by done_date1 desc, wk_item asc"
rstObj1.open strSQL_show,conDB,3
'�Ҧ������u�@�� 
totalputall=rstObj1.recordcount
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�L�n������';background-color:'#F0FFF0'}
input{font-family:'�L�n������';}
textarea{font-family:'�L�n������';}
SELECT{font-family:'�L�n������';font-size:12pt;}
td{font-family:'�L�n������';}
--></style>
</HEAD>
<BODY style="margin-top:5px;">
<center>
<%=worker%>��ثe(<%=date()%>)����w�������u�@���ơG<%=totalputall%>���C<br>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'�}�l�~�� 
ystart=2002 
' 
for yi = ystart to year(date())

dateua=cstr(yi)&"/12/31" 
datena=cstr(yi)&"/1/1" 
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_checker like '%"&worker&"%' and doing_date1 <= #"&dateua &"# and doing_date1 >= #"&datena &"# order by done_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalputa=rstObj1.recordcount

	kkk = (yi-ystart+1) mod 7 
	if kkk=0 then 
		rstr="�C<br>" 
	else
		if yi < year(date()) then rstr="�B"
	end if 
%>
	<a href="wk_lst_done.asp?cyear=<%=yi%>" ><%=yi%>[<%=totalputa%>��]</a><%=rstr%>	 
<% 
	rstr=""

	'������ƶ�
	rstObj1.Close
	'���]����ܼ� 
	set rstObj1=Nothing

next  

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
<hr> 

<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

<%
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_checker like '%"&worker&"%' order by done_date1 desc"
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_checker like '%"&worker&"%' and doing_date1 <= #"&dateu &"# and doing_date1 >= #"&daten &"# order by done_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=5><%=worker%>�G�������b<%=daten%>��<%=dateu%>���A�L�����u�@�ƶ�</font>
<%
else
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=50>
<col width=130>
<col width=130>
<col width=320>
<col width=140>
<tr >
	<td colspan=5 align=center>
	<font size=4><%=worker%>�G�������b<%=daten%>��<%=dateu%>���A�w�����u�@�ƶ��@:<font color=red><%=totalput%></font>��</font>
	</td>
</tr>
<tr >
	<td align=center>�Ǹ�</td>
	<td align=center>������</td>
	<td align=center>�������</td>
	<td align=center>�D��</td>
	<td align=center>
		<a href="./pj_add.asp" target="_blank"> <img src="./img/add1.gif" alt="�M�׷s�W" width="15" height="15" style='cursor:hand;border:0;'></a>
	�M�צW��
		<a href="./pj_list.asp" target="_blank"> <img src="./img/list1.gif" alt="�M�צC��" width="15" height="15" style='cursor:hand;border:0;'></a>
	</td>
</tr>
<%
	'�C�X��ƶ���
	rstobj1.MoveFirst
	for i=1 to totalput
	'Ū�����
		wk_id=rstObj1.fields("wk_id")
		undo_date1=rstObj1.fields("undo_date1")
		doing_date1=rstObj1.fields("doing_date1")
		done_date1=rstObj1.fields("done_date1")
		wk_item=rstObj1.fields("wk_item")
		wk_order=rstObj1.fields("wk_order")
		Response.Write( "<tr>")		
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center style='text-align:right;padding-right:2pt;'><font size=3>" & doing_date1 &" ("&week_cstr(doing_date1)&")</font></td>")
		Response.Write( "<td align=center style='text-align:right;padding-right:2pt;'><font size=3>" & done_date1 &" ("&week_cstr(done_date1)&")</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_order &"</font></td>")
		strA="<a href=wk_show_done.asp?wk_id="& rstObj1.fields("wk_id")&">"
		Response.Write( "<td align=left>" & strA & wk_item &"</a></td>")

		if pj_id="" or isnull(pj_id) then
%>		
		<td align=left><font size=3>
	<a href="./pj_add.asp?wk_id=<%=wk_id%>" target="_blank"> <img src="./img/add1.gif" alt="�s�W�M�צW��" width="15" height="15" style='cursor:hand;border:0;'></a>
	<a href="./pj_sel.asp?wk_id=<%=wk_id%>" target="_blank"> <img src="./img/sel1.gif" alt="��ܱM�צW��" width="15" height="15" style='cursor:hand;border:0;'></a>
		</font></td>
<%
		else
%>		
		<td align=left><font size=3>
	<a href="./pj_delsel.asp?wk_id=<%=wk_id%>&p_id=<%=pj_id%>" target="_blank"> <img src="./img/del1.gif" alt="�����M�צW��" width="15" height="15" style='cursor:hand;border:0;'></a>
	<a href="./pj_show.asp?p_id=<%=pj_id%>" target="_blank"><%=pj_02%></a>
		</font></td>
<%
		end if

		Response.Write( "</tr>")
	'����U�@���O��
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>
</table>
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
<center>
</body>
</html>
