<% @codepage=950%>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

<%
wkgroup="�@��u�@"
	'Ū���H���m�W
	worker = Session("worker")
	'Ū�����Ѥ��
	wk_class=request("wk_class")
	if wk_class="������" then
		wk_class_t="������"
		wk_class=""
	else
		wk_class_t=wk_class
	end if 
	wk_man=request("wk_man")
	if wk_man="�����H��" then
		wk_man_t="�����H��"
		wk_man=""
	else
		wk_man_t=wk_man
	end if 
	if wk_class="" then
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_doer like '%"& wk_man &"%' order by doing_date1 desc"
	else
strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_class like '%"&wk_class&"%' and wk_doer like '%"& wk_man &"%' order by doing_date1 desc"
	end if
%>

<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�L�n������';background-color:'#F0FFF0'}
input{font-family:'�L�n������';font-size:12pt;}
textarea{font-family:'�L�n������';}
SELECT{font-family:'�L�n������';font-size:12pt;}
td{font-family:'�L�n������';}
.tdtext{
		font-size:4mm;
		} 
.tittext{
		font-size:4.5mm;
		font-weight:bold;
		} 
--></style>
</HEAD>
<BODY>
<center>
<font style="font-size:5mm;color:#0000ff;">
�d�߱���G�u�@����=[<font color=red><%=wk_class_t%></font>]�Τu�@�H��=[<font color=red><%=wk_man_t%></font>] 
</font>
<%
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'strSQL_show="Select * from " & tb_name & " where wk_class like '%"&wk_class&"%' and wk_doer like '%"& wk_man &"%' order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4>�d�ߵ��G�G�L�u�@�ƶ�</font>
<%
else
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=50>
<col width=100>
<col width=420>
<col width=200>
<tr >
	<td colspan=4 align=center>
	<font size=4>�d�ߵ��G�G�u�@�ƶ��@��<font color=red><%=totalput%></font>��</font>
	</td>
</tr>
<tr >
	<td align=center>�Ǹ�</td>
	<td align=center>������</td>
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
		wk_gp=trim(rstObj1.fields("wk_group"))
		wk_id=rstObj1.fields("wk_id")
		undo_date1=rstObj1.fields("undo_date1")
		doing_date1=rstObj1.fields("doing_date1")
		wk_item=rstObj1.fields("wk_item")
		wk_order=rstObj1.fields("wk_order")
		pj_id=rstObj1.fields("pj_id")
		pj_02=rstObj1.fields("pj_02")
		Response.Write( "<tr>")
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center><font size=3>" & doing_date1 &"</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_order &"</font></td>")
		strA="<a href=wk_show.asp?wk_id="& rstObj1.fields("wk_id")&">"
		if wk_gp="�@��u�@" then
			Response.Write( "<td align=left>" & strA & wk_item &"</a></td>")
		else
			strA1="<<�M�פu�@>>"
			Response.Write( "<td align=left style='background-color:#ffff99;'>" & strA &strA1 & wk_item &"</a></td>")
		end if
		
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
