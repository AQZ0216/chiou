<% @codepage=950%>
<!-- #Include file = "./include/f_week_cstr.inc" -->
<%
session("hback_URL")="./wk_lst_doing.asp" '���j�T�����^�_����
	'Ū���H���m�W
	worker = Session("worker")
	'Ū�����Ѥ��
	ckdate=date()+2
wkgroup="�@��u�@"

wk_sort=request("wk_sort")
if wk_sort="" then
  wk_sort=session("wk_sort")
   if wk_sort="" then wk_sort=0
end if
session("wk_sort")=wk_sort
   wk_sort_1=1-wk_sort
  ' if wk_sort=0 then
      'wk_sort_1=1
  ' elseif wk_sort=1 then
      'wk_sort_1=0
   'end if
   '�]�wsession("strbackURL")
strbackURL="wk_lst_doing.asp"
session("strbackURL")=strbackURL
%>
<%
'�d�߬O�_������
Function exist_attach(pwk_id)
      ' �s��Access��Ʈwdaywork.mdb
      DBpath_fe=Server.MapPath("./database/attach_file.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '�إ߸�Ʈw�s������
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB_fe.Open strCon_fe
      '�}�Ҹ�ƪ�W��
      tb_name_fe="file_data"
      '�إ߸�Ʈw�s������	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where del_ok = false and wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '������ƶ�
      rstObj1_fe.Close
      '���]����ܼ�
      set rstObj1_fe=Nothing
      '������Ʈw 
      conDB_fe.Close
      '���]�����ܼ�
      set conDB_fe=Nothing
      exist_attach=totalput_fe
End Function

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
'strSQL_show="Select * from " & tb_name & " where wk_doer like '%"&worker&"%' and doing_date1 <= #"&ckdate &"# and wk_undoer like '%"&worker&"%' order by doing_date1 desc"
if  wk_sort=1 then
   strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_doer like '%"&worker&"%' and doing_date1 <= #"&ckdate &"# and doing_date1 >= #"&ckdate-30 &"# and wk_undoer like '%"&worker&"%' order by doing_date1 asc, wk_item asc"
else
   strSQL_show="Select * from " & tb_name & " where wk_group like '%"&wkgroup&"%' and wk_doer like '%"&worker&"%' and doing_date1 <= #"&ckdate &"# and doing_date1 >= #"&ckdate-30 &"# and wk_undoer like '%"&worker&"%' order by doing_date1 desc, wk_item asc"
end if
rstObj1.open strSQL_show,conDB,3
totalput=rstObj1.recordcount
if totalput=0 then
%>
	<font size=4><%=worker%>����(<%=ckdate-2%>)�B����(<%=ckdate-1%>)�Ϋ��(<%=ckdate%>)�ҵL�u�@�ƶ�</font>
<%
else
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=50>
<col width=130>
<col width=420>
<col width=30>
<col width=140>
<tr >
	<td colspan=5 align=center>
	<font size=4><%=worker%>����(<%=ckdate-2%>)�B����(<%=ckdate-1%>)�Ϋ��(<%=ckdate%>)���u�@�ƶ��@:<font color=red><%=totalput%></font>��</font>
	</td>
</tr>
<tr >
	<td align=center>�Ǹ�</td>
	<td align=center><a href="wk_lst_doing.asp?wk_sort=<%=wk_sort_1%>" >������</a></td>
	<td align=center>�D��</td>
	<td align=center>
	<a href="0_wk_headline_off_all.asp?wk_id=<%=wk_id%>" style="text-decoration:none;color:green;" title="�M���Ҧ����j�T���I�I">��</a>
	</td>
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
            wk_item=replace(wk_item,"���׶�","<font color=fuchsia  >���׶�</font>")
		wk_order=rstObj1.fields("wk_order")
		pj_id=rstObj1.fields("pj_id")
		pj_02=rstObj1.fields("pj_02")
		p_headline=rstObj1.fields("headline")
		'�ˬd�O�_������ exist_attach(wk_id)
               attach_no=exist_attach(wk_id)
               if attach_no=0 then
                  str_colors="color:#000000;"
               else
                  str_colors="color:#0000FF;"
               end if
               if rstObj1.fields("wk_password")="" or isnull(rstObj1.fields("wk_password")) then
               else
                  str_colors="color:#0000FF;"
               end if
		Response.Write( "<tr>")
		Response.Write( "<td align=center><font size=3>" & i &"</font></td>")		
		Response.Write( "<td align=center style='text-align:right;padding-right:2pt;"& str_colors &"'><font size=3>" & doing_date1 &" ("&week_cstr(doing_date1)&")</font></td>")
		'Response.Write( "<td align=center><font size=3>" & wk_order &"</font></td>")
		strA="<a href=wk_show.asp?wk_id="& rstObj1.fields("wk_id")&" style='letter-spacing:1.5pt;font-size:12pt;' >"
		if wk_gp="�@��u�@" then
			Response.Write( "<td align=left >" & strA & wk_item &"</a></td>")
		else
			strA1="<<�M�פu�@>>"
			Response.Write( "<td align=left style='background-color:#ffff99;'>" & strA &strA1 & wk_item &"</a></td>")
		end if
%>
      <td align=center>
<% if p_headline<= 5 or isnull(p_headline) then%>
   <a href="0_wk_headline_on.asp?wk_id=<%=wk_id%>" style="text-decoration:none;color:black;" title="�C�����j�T���I�I">��</a>
<% else %>
   <a href="0_wk_headline_off.asp?wk_id=<%=wk_id%>" style="text-decoration:none;color:red;" title="�������j�T���I�I">��</a>
<% end if %>
      </td>
<%		
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
