<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
<!-- �}�Ҹ�Ʈw -->
<!-- Include file = "./include/opendb_daywork.inc" -->
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
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'Ū�����
undo_date1=rstObj1.fields("undo_date1")
doing_date1=rstObj1.fields("doing_date1")
done_date1=rstObj1.fields("done_date1")
wk_item=rstObj1.fields("wk_item")
wk_content=rstObj1.fields("wk_content")
wk_order=rstObj1.fields("wk_order")
wk_doer=rstObj1.fields("wk_doer")
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_class=rstObj1.fields("wk_class")
wk_group=rstObj1.fields("wk_group")
wk_exe=rstObj1.fields("wk_exe")
wk_pjn=rstObj1.fields("pj_02")   '�M�צW��
wk_att=rstObj1.fields("wk_att")           '�X�u�H��

'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
<%
' �s��Access��Ʈwtemp-daywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="work_data"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where tmp_id ="&wk_id&" and ipt_ok=0 order by wk_id desc" 
rstObj1.open strSQL_show,conDB,1,1
tpn=rstObj1.recordcount
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
body{font-family:'�з���';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<form name="form1" action="" method="post">
<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<input type="hidden" name="worker1" value="<%=worker%>">
<input type="hidden" name="wk_order1" value="<%=wk_order%>">
<center>

<table border=0 cellspacing=0 cellpadding=0>
<col span=7 style="width:100px;text-align:center;">
<tr width=720>	
<% if tpn=1 then %>
<td>�i���P�B�j</td>
<% else %>
<td>�i�w�P�B�j</td>
<% end if %>
	<td>
		<input type=button name="bkpg" value="�^�W�@��" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="parent.location.href='javascript:history.back()'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
<% if tpn=1 then %>
	<td> 	<input type=button name="edit" value="�s�ץ��P�B�u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="wk_edit()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';"> 	</td>
	<td>	<input type=button name="delete" value="�R�����P�B�u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="wk_del()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">	</td>
<% else %>
	<td>	<input type=button name="delete" value="�R���w�P�B�u�@" title="�R���w��s���u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="wk_delnext()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">	</td>
<% end if %>
	<td>	<input type=button name="wkprint" value="�^����" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="wk_calendar('<%=year(doing_date1)%>','<%=month(doing_date1)%>')"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';"></td>

</tr>
</table>
<table border=1 cellspacing=0 cellpadding=0>
<col width=120><col width=120><col width=120><col width=120><col width=120><col width=120>
<tr>
	<td align="center" colspan=2><font size=4 color="red"><b>��ܳ�@�u�@��(<%=wk_group%>)</b></font></td>
	<td align="right">�u�@�s���G</td>
	<td><%=wk_id%></td>
	<td align="right">�u�@�����G</td>
	<td><%=wk_class%></td>

</tr>
<tr>
	<td align="right">���i�̡G</td>
	<td><%=wk_order%></td>
	<td align="right">���i����G</td>
	<td><%=undo_date1%></td>
	<td align="right">�������G</td>
	<td><%=doing_date1%></td>
</tr>

<tr>
	<td align="right">
	���|�H���G
	</td>
	<td colspan=5>
	<%=wk_doer%>
	</td>
</tr>
<tr>
	<td align="right">
	����H���G
	</td>
	<td colspan=5>
	<%=wk_exe%>
	</td>
</tr>
<tr>
	<td align="right">
	�X�u�H���G
	</td>
	<td colspan=5>
	<%=wk_att%>
	</td>
</tr>
<!--
<tr>
	<td align="right">
	�����H���G
	</td>
	<td colspan=5>
	<%=wk_checker%>
	</td>
</tr>
-->
<!--<tr>
	<td align="right">
	�������H���G
	</td>
	<td colspan=5>
	<%=wk_undoer%>
	</td>
</tr>-->
<tr>
	<td align="right">
	�D���G
	</td>
	<td colspan=5>
	<%=wk_item%>
	</td>
</tr>
<tr>
	<td align="right" valign="top">
	���椺�e�G
	</td>
	<td colspan=5>
	<%
	wk_content_s=replace(wk_content,chr(13),"<br>")
	%>
	<font style="font-size:12pt;" ><%=wk_content_s%></font>
	</td>
</tr>

</table>
</center>

<!-- Include file = "./include/wk_show_form.inc" -->

<script type="text/javascript">
function wk_edit()
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./3_mobilejs_wk_edit.asp?wk_id="+x1;
   // alert(str_url);
   location.href = str_url ;
}
function wk_del()
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./3_mobilejs_wk_del_ok.asp?wk_id="+x1;
		var r = confirm("�нT�{�O�_�R���i���P�B�j���u�@�H�H"+String.fromCharCode(13,10)+str_url);
		if (r == true) {
		    //txt = "�T�{�R���I�I";
		    //alert(txt);
		    location.href = str_url ;
		} else {
		    //txt = "�����R���I�I";
		    //alert(txt);
		}
   // alert(str_url);
   //location.href = str_url ;
}
function wk_calendar(pyear,pmonth)
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./wk_calendar_all.asp?nYear="+pyear+"&nMonth="+pmonth;
   // alert(str_url);
   location.href = str_url ;
}
function wk_delnext()
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./3_mobilejs_wk_delnext_ok.asp?wk_id="+x1;
		var r = confirm("�нT�{�O�_�R���i�w�P�B�j���u�@�H�H"+String.fromCharCode(13)+str_url);
		if (r == true) {
		    //txt = "�T�{�R���I�I";
		    //alert(txt);
		    location.href = str_url ;
		} else {
		    //txt = "�����R���I�I";
		    //alert(txt);
		}
   // alert(str_url);
   //location.href = str_url ;
}
</script>
</form>
</center>

</body>
</html>
