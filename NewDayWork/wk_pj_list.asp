<% @codepage=950%>
<!-- #Include file = "./include/f_week_cstr.inc" -->
<%
	'Ū���H���m�W
	worker = Session("worker")
	'Ū�����Ѥ��
	ckdate=date()+2
wkgroup="�M�פu�@"
%>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->

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
<BODY>
<center>

<%
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
'strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_doer like '%"&worker&"%' and doing_date1 <= #"&ckdate &"# and wk_undoer like '%"&worker&"%' order by doing_date1 desc"
'strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' order by doing_date1 desc"
strSQL_show="Select * from " & tb_name & " where not(isnull(pj_02) or trim(pj_02) like '' or pj_02 like '����|Ū�ѷ|���') or (wk_group like '%"&wkgroup&"%') order by doing_date1 desc"
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4>�L�M�פu�@�ƶ�</font>
<%
else
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=50>
<col width=130>
<col width=420>
<col width=170>
<tr >
	<td colspan=4 align=center>
	<font size=4>�Ҧ��M�פu�@�ƶ��@:<font color=red><%=totalput%></font>��</font>
	</td>
</tr>
<tr >
	<td align=center>�Ǹ�</td>
	<td align=center>������</td>
<!--	<td align=center>���i��</td>-->
	<td align=center>�D��</td>
	<td align=center>
		<a href="./pj_add.asp" target="_blank"> <img src="./img/add1.gif" alt="�M�׷s�W" width="15" height="15" style='cursor:hand;border:0;'></a>
	�M�צW��
		<a href="./pj_list.asp" target="_blank"> <img src="./img/list1.gif" alt="�M�צW�٦C��" width="15" height="15" style='cursor:hand;border:0;'></a>
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
		wk_item=rstObj1.fields("wk_item")
		wk_order=rstObj1.fields("wk_order")
		pj_id=rstObj1.fields("pj_id")
		pj_02=rstObj1.fields("pj_02")
		Response.Write( "<tr>")		
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center style='text-align:right;padding-right:2pt;'><font size=3>" & doing_date1 &" ("&week_cstr(doing_date1)&")</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_order &"</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_exe &"</font></td>")
		strA="<a href=wk_pj_show.asp?wk_id="& rstObj1.fields("wk_id")&">"
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
