<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<!-- Include file = "./include/array_worker.inc" -->
<!-- #Include file = "./include/array_wkclass.inc" -->
<!-- #Include file = "./include/array_wkgroup.inc" -->
<!-- #Include file = "./include/workinput.inc" -->
<!-- #Include file = "./misc_data/array_place.inc" -->	
<!-- #Include file = "./misc_data/array_thing.inc" -->	
<!-- Include file = "./misc_data/array_writer.inc" -->	
<!-- #Include file = "./include/array_pjn.inc" -->
<%
'response.write "�`�g�z��="&stra_dp01&"�C<br>"
'response.write "�޲z��="&stra_dp02&"�C<br>"
'response.write "������="&stra_dp03&"�C<br>"
'response.write "�~�ȳ�="&stra_dp04&"�C<br>"
'response.write "�k�ȳ�="&stra_dp05&"�C<br>"
'response.write "�]�ȳ�="&stra_dp06&"�C<br>"
'response.write "��T��="&stra_dp07&"�C<br>"
'response.write "�س]��="&stra_dp08&"�C<br>"
'response.write "����|="&stra_dp09&"�C<br>"
'response.end
%>
<%
stra_gp6=""
stra_gp5=""
'�H���W��
	'stra_gp5 ���~�H��
'	st_dp5="�]�ȳ�,�k�ȳ�,��T��,�޲z��"
	'stra_gp6 �~�ȤH��
'	st_dp6="������,�~�ȳ�"
'	st_dp7="����|"
for ki=1 to worker_no
	if staff_gp_a(ki-1)="���~" then
		stra_gp5= stra_gp5 & "," & worker_a(ki-1)
	elseif left(staff_gp_a(ki-1),1)="�~" then
		stra_gp6= stra_gp6 & "," & worker_a(ki-1)
'	elseif staff_gp_a(ki-1)="����" then
'		stra_gp7= stra_gp7 & "," & worker_a(ki-1)
	end if
	stra_gp0= stra_gp0 & "," & worker_a(ki-1)
next
'response.write "worker_no="&worker_no&"�C<br>"
'response.write "stra_gp6="&stra_gp6&"�C<br>"
'response.write "stra_gp5="&stra_gp5&"�C<br>"
'response.write "stra_gp7="&stra_gp7&"�C<br>"
'response.end
stra_gp0=right(stra_gp0,len(stra_gp0)-1) '����
stra_gp6=right(stra_gp6,len(stra_gp6)-1) '�~�ȤH��
stra_gp5=right(stra_gp5,len(stra_gp5)-1) '���~�H��
'stra_gp7=right(stra_gp7,len(stra_gp7)-1) '����
stra_gp1="����,���,���,����,����,���z,�f�S,���F,�_��"   '������{�M��
%>
<%
	'Ū���H���m�W
	worker = Session("worker")
	datecode=request("datecode")
	'if datecode="" then datecode=date()
	wk_order=worker
	undo_date1=date()
'�u�@���Ű}�C 
'dim wk_class_a
'wk_class_a=array("","A","B","C","D")
'wk_class_no=ubound(wk_class_a)+1
'�u�@���Ű}�C 
'dim wk_group_a
'wk_group_a=array("�@��u�@","�M�פu�@")
'wk_group_no=ubound(wk_group_a)+1
%>


' <%
' '�P�_�O�_�OIE�Τ��
' dim u,b
' set u=Request.ServerVariables("HTTP_USER_AGENT")
' 'response.write u
' 'response.write "<hr>"
' set b=new RegExp
' b.Pattern="firefox|chrome|safari|mobile"
' b.IgnoreCase=true
' b.Global=true
' Set matchesb = b.Execute(u)
' if b.test(u) then               '�DIE�s����
'       'response.redirect("http://detectmobilebrowser.com/mobile")
'       'response.write "b="& matchesb(0).value &"<hr>"
'       'response.write "b.test(u)="&b.test(u)&"<hr>"
'       'response.write "�s�����G"& matchesb(0).value & "<hr>"
'       '�DIE
'       nexturl="3_mobilejs_wk_add.asp?datecode="&datecode
'       response.redirect(nexturl)
' else
'       'response.write "b.test(u)="&b.test(u)&"<hr>"
'       'response.write "�s�����G"&"IE<hr>"
' end if
' %>
<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body {  scrollbar-3dlight-color:#ffffff;
        scrollbar-arrow-color:#CCCCCC;
        scrollbar-base-color:#666633;
        scrollbar-darkshadow-color:#e6e6cc;
        scrollbar-face-color:#666666;
        scrollbar-highlight-color:#ffffff;
        scrollbar-shadow-color:#e6e6cc;
        scrollbar-track-color:#ffffff;
        margin:2mm 0mm 0mm 0mm;		/*��t�W�U���k*/
		font-family:'�L�n������';		/*�r��*/
		font-size:4mm; 			/*�r��j�p*/
		background-color:'#F0FFF0';
     }
input.imenu { 
	font-size:3.5mm;				/*�r��j�p*/
	cursor:hand;				/*��ЧΦ�*/ 
	background-color:'#d3d3d3'; 		/*�~���C��*/ 
	margin:0 0 0 0;		/*��t�W�U���k*/
	width:40px;
     }
input.imenu1 { 
	font-size:3.5mm;	/*�r��j�p*/
	font-weight:bold;				
	cursor:hand;				/*��ЧΦ�*/ 
	background-color:'#eeeeff'; 		/*�~���C��*/ 
	margin:0 0 0 0;		/*��t�W�U���k*/
	width:80px;
	height:100%;
     }
     
TD.SOME{
		font-family: '�L�n������';
		font-size: 3.3mm;
		line-height: 18px;
		color:blue;
		font-weight:bold;
		}
TD.myd{
		font-family: '�L�n������';
		font-size: 3.3mm;
		line-height: 18px;
		background-color:#f0ffff;
		}     
    
-->
</style>

</HEAD>
<BODY>
<center>

<form name="form1" action="wk_add_ok.asp" method="post" >
<input type="hidden" name="worker1" value="<%=worker%>" >
<input type="hidden" name="headline1" value="0" >
<table border=1 cellspacing=0 cellpadding=0>
<col width=120>
<col width=180>
<col width=100>
<col width=180>
<col width=100>
<col width=210>
<tr>
	<td colspan=2 align=center><font size=4 color="red"><b><%=worker%>�u�@���i��</b></font></td>
	<td align="right"><font color="red">�u�@�s�աG</font></td>
	<td>
	<select name="wk_group" style="width:100%;height:100%;font-size:10pt;background-color:'#ffffee';">
<%
		response.write "<option value='"&wk_group_a(0)&"' selected>"&wk_group_a(0)
	for i=2 to wk_group_no
		response.write "<option value='"&wk_group_a(i-1)&"'>"&wk_group_a(i-1)
	next
%>
	</select>
	</td>
	<td align="right">
	<font color="red"><a href="./pj_add.asp" target="_self" title="�s�W�M�צW��"> �M�צW��</a>�G</font>
	</td>
	<td>
	<select name="wk_pjn" style="width:100%;height:100%;font-size:10pt;background-color:'#ffffee';" >
<%
		response.write "<option value='0' selected>"
		'response.write "<option value='"&pjnid_a(0)&"' >"&pjn_a(0)
	for i=1 to pjn_no
		response.write "<option value='"&pjnid_a(i-1)&"�A"&pjn_a(i-1)&"'>"&pjn_a(i-1)
	next
%>
	</select>
	</td>	
<tr>
	<td align="right">���i�̡G</td>
	<td><input type='text' name='wk_order' value='<%=wk_order%>' style="width:100%;" readonly></td>
	<td align="right">���i����G</td>
	<td><input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly></td>
	<td align="right"><font color="red">�u�@�����G</font></td>
	<td>
	<select name="wk_class" style="width:100%;height:100%;font-size:10pt;background-color:'#ffffee';" onchange="item_chk()">
<%
	for i=1 to wk_class_no
		response.write "<option value='"&wk_class_a(i-1)&"'>"&wk_class_a(i-1)
		 if wk_class_a(i-1)="Z" then response.write "-���n����"
	next
%>
	</select>
	</td>
</tr>
<tr>
	<td colspan=6 align="center">
<table border="0" cellspacing="0" cellpadding="0">
<col width=120><col width=120><col width=120><col width=120><col width=120>
<tr><td align="center" valign="middle">
	<!-- �H���ﶵ -->
	<font size=3>�H���ﶵ�G<br>
		<SELECT name="men_w" onchange="menadd()">
		<option selected>�п�ܤH��</option>
	<%
		for i=1 to worker_no
			response.write "<option value='"&worker_a(i-1)&"'>"&worker_a(i-1)
		next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- ����ﶵ -->
	<font size=3>����ﶵ�G<br>
	<input type='hidden' name="doing_date2" value="" style="width:70%;" onchange="dateadd()">
		<img align=top src="img/cal3.gif" onmousedown="Cal('doing_date2')" width="20" height="20" style='cursor:hand;'>
	</td>
	<td align="center" valign="middle">
	<!-- �ɶ��ﶵ -->
	<font size=3 color=red>����ɶ��G<br>
		<SELECT name="time_w" onchange="timeadd()">
		<option value="" selected>�п�ܮɶ�</option>
	<%
	for i=1 to 48
		Response.Write("<OPTION value=" & wk_time(i-1) & ">" & wk_time(i-1)&"</OPTION>")
	next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- �a�I�ﶵ -->
	<font size=3>�a�I�ﶵ�G<br>
		<SELECT name="place_w" onchange="placeadd()">
		<option selected>�п�ܦa�I</option>
	<%
		for i=1 to place_no
			response.write "<option value='"&place_a(i-1)&"'>"&place_a(i-1)
		next
	%>
		</font></SELECT>
	</td>
	<td align="center" valign="middle">
	<!-- �ƥ�ﶵ -->
	<font size=3>�ƥ�ﶵ�G<br>
		<SELECT name="thing_w" onchange="thingadd()">
		<option selected>�п�ܨƥ�</option>
	<%
		for i=1 to thing_no
			response.write "<option value='"&thing_a(i-1)&"'>"&thing_a(i-1)
		next
	%>
		</font></SELECT>
	</td>

</tr>
</table>
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">�D���G</font>
	</td>
	<td colspan=3>
	<input type='text' name='wk_item' value='' style="width:100%;" onchange="item_chk()">
	</td>
	<td align="right"><font color="red">�������G</font></td>
	<td><input type='text' name="doing_date1" value="<%=datecode%>" style="width:70%;">
		<img align=top src="img/cal3.gif" onmousedown="Cal('doing_date1')" width="20" height="20" style='cursor:hand;'>
	</td>

</tr>
<tr style="background-color:#FFFF33;">
	<td align="right">
		<font style="background-color:#ddd;text-decoration:none;cursor:hand;color:red;" onclick="addexe_none()" title="�M������H�����">[�M]</font>
	<font style="color:blue;">����H��</font>
	</td>
	<td colspan=5>
	<input type='text' name='wk_exe' value='' style="width:50%;" readonly title="����H���бĥΥk��U�Կ���J�I�I�I" onkeydown="javascript:if(window.event.keyCode==8) return false;">
		<SELECT name="exemen_w" onchange="exeadd()">
		<option value="" selected>�п�ܤH��</option>
		<option value="clear" >�M���H��</option>
			<option value="����H��" >����H��</option>
		<option value="�~�ȥ���" >�~�ȥ���</option>
		<option value="���ԥ���" >���ԥ���</option>
	<%
		for i=1 to worker_no
			response.write "<option value='" & worker_a(i-1) & "'>" & worker_a(i-1) &"</option>"
		next
	%>

		</SELECT>

		<SELECT name="exemen_dp" onchange="exeadddp()">
			<option value="" selected>�������</option>
			<option value="clear" >�M���H��</option>
			<option value="<%=stra_dp01%>" >�`�g�z��</option>
			<option value="<%=stra_dp02%>" >�޲z��</option>
			<option value="<%=stra_dp03%>" >������</option>
			<option value="<%=stra_dp04%>" >�~�ȳ�</option>
			<option value="<%=stra_dp05%>" >�k�ȳ�</option>
			<option value="<%=stra_dp06%>" >�]�ȳ�</option>
			<option value="<%=stra_dp07%>" >��T��</option>
			<option value="<%=stra_dp08%>" >�س]��</option>
			<option value="<%=stra_dpa1%>" >�~1</option>
			<option value="<%=stra_dpa2%>" >�~2</option>
			<option value="<%=stra_dpa3%>" >�~3</option>
		</SELECT>	
				(�п�J����ѻP�H��)	
	</td>
</tr>

<tr style="background-color:#FFBFFF;">
	<td align="right">
		<font style="background-color:#ddd;text-decoration:none;cursor:hand;" onclick="addsameatt_exe()" title="�P����H��">[�P]</font>		
		<font style="background-color:#ddd;text-decoration:none;cursor:hand;color:red;" onclick="addatt_none()" title="�M���X�u�H�����">[�M]</font>
	<font style="color:blue;">�X�u</font>
	</td>
	<td colspan=5>
	<input type='text' name='wk_att' value='' style="width:50%;" readonly title="�X�u�H���бĥΥk��U�Կ���J�I�I�I" >
		<SELECT name="attmen_w" onchange="attadd()">
		<option value="" selected>�п�ܤH��</option>
		<option value="clear" >�M���H��</option>
			<option value="����H��" >����H��</option>
		<option value="�~�ȥ���" >�~�ȥ���</option>
		<option value="���ԥ���" >���ԥ���</option>
	<%
		for i=1 to worker_no
			response.write "<option value='" & worker_a(i-1) & "'>" & worker_a(i-1) &"</option>"
		next
	%>

		</SELECT>

		<SELECT name="attmen_dp" onchange="attadddp()">
			<option value="" selected>�������</option>
			<option value="clear" >�M���H��</option>
			<option value="<%=stra_dp01%>" >�`�g�z��</option>
			<option value="<%=stra_dp02%>" >�޲z��</option>
			<option value="<%=stra_dp03%>" >������</option>
			<option value="<%=stra_dp04%>" >�~�ȳ�</option>
			<option value="<%=stra_dp05%>" >�k�ȳ�</option>
			<option value="<%=stra_dp06%>" >�]�ȳ�</option>
			<option value="<%=stra_dp07%>" >��T��</option>
			<option value="<%=stra_dp08%>" >�س]��</option>
			<option value="<%=stra_dpa1%>" >�~1</option>
			<option value="<%=stra_dpa2%>" >�~2</option>
			<option value="<%=stra_dpa3%>" >�~3</option>
		</SELECT>	
				(�п�J�����X�u�H��)	
	</td>
</tr>

<tr>
	<td align="right">
	<font color="red">���椺�e�G</font>
	</td>
	<td colspan=5>
	<TEXTAREA name="wk_content" rows="5" style="width:100%;" ><%=wk_content%></TEXTAREA>
	</td>
</tr>

<tr>
	<td align="right">
				<font style="background-color:#ddd;text-decoration:none;cursor:hand;color:red;" onclick="addaw_none()" title="�M�����|�H�����">[�M]</font>
	<font color="red">���|�H���G</font>
		<font style="background-color:#ddd;text-decoration:none;cursor:hand;" onclick="addsame_exe()" title="�P����H��">[�P����H]</font>		

	</td>
	<td colspan=5>
	<input type='text' name='all_worker' value='<%=worker%>' onchange="item_chk()" readonly style="width:100%;" onkeydown="javascript:if(window.event.keyCode==8) return false;">

		<font style="background-color:#ddd;text-decoration:none;cursor:hand;" onclick="addaw_all()" title="����H���q��">[����]</font>
		<font style="background-color:#ddd;text-decoration:none;cursor:hand;" onclick="addaw_gp6()" title="�~�ȤH��">[�~��]</font>

		<SELECT name="wkrmen_w" onchange="wkradd()">
		<option value="" selected>�п�ܤH��</option>
		<option value="clear" >�M���H��</option>
			<!--<option value="����H��" >����H��</option>
		<option value="�~�ȥ���" >�~�ȥ���</option>-->
	<%
		for i=1 to worker_no
			response.write "<option value='" & worker_a(i-1) & "'>" & worker_a(i-1) &"</option>"
		next
	%>
		</SELECT>
			<SELECT name="wkrmen_dp" onchange="wkradddp()" style="vertical-align:top;">
			<option value="" selected>�������</option>
			<option value="clear" >�M���H��</option>
			<option value="" selected>�������</option>
			<option value="<%=stra_dp01%>" >�`�g�z��</option>
			<option value="<%=stra_dp02%>" >�޲z��</option>
			<option value="<%=stra_dp03%>" >������</option>
			<option value="<%=stra_dp04%>" >�~�ȳ�</option>
			<option value="<%=stra_dp05%>" >�k�ȳ�</option>
			<option value="<%=stra_dp06%>" >�]�ȳ�</option>
			<option value="<%=stra_dp07%>" >��T��</option>
			<option value="<%=stra_dp08%>" >�س]��</option>
<!--			<option value="<%=stra_dp09%>" >����</option>-->
			<option value="<%=stra_dp10%>" >�ڮa�A�~</option>
			<option value="<%=stra_dpa1%>" >�~�@</option>
			<option value="<%=stra_dpa2%>" >�~�G</option>
			<option value="<%=stra_dpa3%>" >�~Three</option>
			<option value="<%=stra_dpa4%>" >YES</option>
			<option value="<%=stra_dpa5%>" >�_�w�~�K</option>

		</SELECT>	

	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">�q���y���G</font>
	</td>
	<td colspan=5>
	<input type='radio' name='golf_ok' value='�O' >�O�C
	<input type='radio' name='golf_ok' value='�_' checked>�_�C
	<font color=red >(�q���y���ȯ�榸���i�C)</font>
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">���Ƥ��i�G</font>
	</td>
	<td colspan=5>
	   <select name="repeat_type" >
	     <option value="" >�п�ܡK</option>
	     <option value="cs_1" selected>�榸</option>
	     <option value="cs_week1" >�C�P�@��</option>
	     <option value="cs_week2" >��P�@��</option>
	     <option value="cs_month1" >�C��@��</option>
	     <option value="cs_year1" >�C�~�@��</option>
	     <option value="cs_m_first_monday" >�C�몺�Ĥ@�ӬP���@</option>
	     <option value="cs_manual" >�ۭq���</option>
	   </select>	
	   �C
	   �I�����(���t)�G<input type='text' name="end_date" value='<% = date() +1%>' style="width:90px;text-align:right;">
	     <img align=top src="img/cal3.gif" onmousedown="Cal('end_date')" width="20" height="20" style='cursor:hand;'>
	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">�ۭq����G</font>
	</td>
	<td colspan=5>
	<TEXTAREA name="rp_dates" rows="3" style="width:100%;" ></TEXTAREA>
	����H�r��(,)���j�A�d�ҡG2011/01/01,2011/03/02,2011/05/03
	</td>
</tr>
<tr>
	<td colspan=6 align="center">
	<input type="button" name="press" value="�T�w���i" onclick="press_chk()">
	<input type="reset" name="cancel" value="�M�����" >
	</td>
<tr>
<tr>
	<td align="right">
	<font color="red">�[�K��r�G</font>
	</td>
	<td colspan=5>
	<input type='text' name='str_pwd' value='' style="width:100px;" maxlength=10>
	</td>
</tr>
</table>
</form>

<!--��䲣�ͪ���m-->
<Span ID=ShowCal style="position:absolute;z-index:1;"></Span>

<Script Language=VBScript>
<!--
Sub Cal(TObject)

'���ͤ��A�H���Ѫ���������
 Call GetCal(Year(Now()),Month(Now()),TObject)

'�վ�<Span>����m
 ShowCal.style.left=window.event.clientX-140
 ShowCal.style.top=window.event.clientY+10
End Sub 

Sub GetCal(SYear,SMonth,TObject)
'�����Y����
Str=Str &"<table ALIGN='CENTER' BORDER='0' CELLSPACING='0' CELLPADDING='2' BGCOLOR='#f0ffff' BORDERCOLOR='Gray'>"
Str=Str &"<tr><td>"
Str=Str &"        <table WIDTH='140' BORDER='0' CELLPADDING='1' CELLSPACING='0' BGCOLOR='#FFFFFF'>"
Str=Str &"                <tr HEIGHT='18' BGCOLOR='Silver'>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/prev3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='�W�@��' style='cursor:hand' OnClick='PreMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/Next3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='�U�@��' style='cursor:hand' OnClick='NextMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='100' COLSPAN='4' ALIGN='CENTER' VALIGN='MIDDLE' CLASS='SOME'><Span ID=SelYear>" & SYear & "</span>�~<Span ID=SelMon>" & SMonth & "</Span>��</td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='RIGHT' VALIGN='MIDDLE'><img SRC='img/cdia3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='��������' style='cursor:hand;' OnClick='Closedate()' ></td>"
Str=Str &"                </tr>"
Str=Str &"          <tr HEIGHT='15' BGCOLOR='Aliceblue'>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>��</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>�@</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>�G</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>�T</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>�|</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>��</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>��</td>"
Str=Str &"          </tr>"
Str=Str &"      <tr>"

'�Ӥ�Ĥ@�Ѥ��P��
SDate=DateValue(SYear & "/" & SMonth & "/1")
SWeek=WeekDay(SDate)

'�Ӥ�̫�@�Ѫ����
EDay=Day(DateSerial(SYear,SMonth+1,0))
EDate=DateValue(SYear & "/" & SMonth & "/" & EDay)

'�Ӥ�̫�@�Ѥ��P��
EWeek=WeekDay(EDate)

'�W��̫�@�Ѫ����
PreEDay=Day(DateSerial(SYear,SMonth,0))

'���ͤW�Ӥ몺����
Dim i
For i=1 to SWeek-1
 Str=Str & "<td CLASS='myd' width=20 align=right><font color=gray>" & PreEDay-SWeek+i+1 & "</font></td>"
Next

If SWeek=1 then
 Str=Str & "</tr>"
End if

'���ͥ��볡��
SW=SWeek
i=1
For i=1 to EDay
 '�վ㤻��r�����C��
 Select Case SW
 Case 1
  FColor="Red"
 Case 7
  FColor="Green"
 Case Else
  FColor="Black"
 End Select
  
  GDate=SYear & "/" & SMonth & "/" & i
  str_cc=""
  if datevalue(GDate)=date then str_cc="background-color:#99cc99;"
  Str=Str & "<td CLASS='myd' width=20 align=right style='cursor:hand;' onMouseOver=""this.style.backgroundColor='#FF99FF'"" onMouseOut=""this.style.backgroundColor='#f0ffff'"" Onclick=""SendDate('" & GDate & "','" & TObject & "')""><Font style='"&str_cc&"' Color=" & FColor & ">" & i & "</Font></td>"
 
'���ͤU�Ӥ볡��
SW=SW+1
 IF SW>7 then
  Str=Str & "</tr><tr>"
  SW=1
 End if
Next

J=1
For i=SW to 7
 Str=Str & "<td CLASS='myd' width=20 align=right><Font Color=Gray>" & j & "</Font></td>"
 J=j+1
Next


Str=Str & "      </tr>"
Str=Str & "</Table>"
'�N��ƤޤJ<Span>
ShowCal.InnerHTML=Str

End Sub

'�e���@�Ӥ�
Sub PreMon(TObject)
 SYear=Int(SelYear.OuterTEXT)
 SMon=int(SelMon.outerTEXT)-1

 '�P�_�O�_���e�դ@�~
 IF SMon<1 then
  SMon=12
  SYear=SYear-1
 End if
 Call GetCal(SYear,SMon,TObject)
End Sub

'�Ჾ�@�Ӥ�
Sub NextMon(TObject)
 SYear=Int(SelYear.OuterTEXT)
 SMon=int(SelMon.outerTEXT)+1

 '�P�_�O�_���e���@�~
 IF SMon>12 then
  SMon=1
  SYear=SYear+1
 End if
 Call GetCal(SYear,SMon,TObject)
End Sub

'�N��ưe�J��줺
Sub SendDate(GDate,TObject)
 document.all.namedItem(TObject).Value=GDate
 if TObject="doing_date1" or TObject="end_date" then
 else
 	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.all.namedItem(TObject).Value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.all.namedItem(TObject).Value
	end if
end if
 ShowCal.InnerHTML=""
End Sub

'��������e��
Sub CloseDate()
 ShowCal.InnerHTML=""
End Sub
-->
</script>

<script language=vbscript>
<%
for i=1 TO worker_no
%>
sub worker_se<%=i%>_click()
	if document.form1.wk_exe.value="" then
		document.form1.wk_exe.value=Trim(document.form1.worker_se<%=i%>.value)
	else
		document.form1.wk_exe.value=document.form1.wk_exe.value &","& Trim(document.form1.worker_se<%=i%>.value)
	end if
end sub
<%
next
%>
sub all_sele_click()
	document.form1.wk_exe.value=""
	<%
	for i=1 TO worker_no
	%>	
		worker_se<%=i%>_click
	<%
	next
	%>	
end sub
sub all_unsele_click()
	document.form1.wk_exe.value=""
end sub
</script>

<script language=vbscript>
<%
for i=1 TO worker_no
%>
sub worker_s<%=i%>_click()
	if document.form1.all_worker.value="" then
		document.form1.all_worker.value=Trim(document.form1.worker_s<%=i%>.value)
	else
		document.form1.all_worker.value=document.form1.all_worker.value &","& Trim(document.form1.worker_s<%=i%>.value)
	end if
end sub
<%
next
%>
sub all_sel_click()
	document.form1.all_worker.value=""
	<%
	for i=1 TO worker_no
	%>	
		worker_s<%=i%>_click
	<%
	next
	%>	
end sub
sub all_unsel_click()
	document.form1.all_worker.value=document.form1.worker1.value
end sub
</script>
<script language=vbscript>
sub attadd()'�X�u�H��
  if document.form1.attmen_w.value="clear" then
   document.form1.wk_att.value=""
  else
	if document.form1.wk_att.value="" then
		document.form1.wk_att.value=document.form1.attmen_w.value
	else
         if instr(1,document.form1.wk_att.value,document.form1.attmen_w.value,1)>0 then
            document.form1.wk_att.value=replace(document.form1.wk_att.value,document.form1.attmen_w.value,"")
            document.form1.wk_att.value=replace(document.form1.wk_att.value,",,",",")
            if left(document.form1.wk_att.value,1)="," then document.form1.wk_att.value=right(document.form1.wk_att.value,len(document.form1.wk_att.value)-1)
            if right(document.form1.wk_att.value,1)="," then document.form1.wk_att.value=left(document.form1.wk_att.value,len(document.form1.wk_att.value)-1)
         else
		document.form1.wk_att.value=document.form1.wk_att.value & "," & document.form1.attmen_w.value
         end if
	end if
  end if
	document.form1.attmen_w.value=""
end sub	
sub exeadd()
  if document.form1.exemen_w.value="clear" then
   document.form1.wk_exe.value=""
  else
	if document.form1.wk_exe.value="" then
		document.form1.wk_exe.value=document.form1.exemen_w.value
	else
         if instr(1,document.form1.wk_exe.value,document.form1.exemen_w.value,1)>0 then
            document.form1.wk_exe.value=replace(document.form1.wk_exe.value,document.form1.exemen_w.value,"")
            document.form1.wk_exe.value=replace(document.form1.wk_exe.value,",,",",")
            if left(document.form1.wk_exe.value,1)="," then document.form1.wk_exe.value=right(document.form1.wk_exe.value,len(document.form1.wk_exe.value)-1)
            if right(document.form1.wk_exe.value,1)="," then document.form1.wk_exe.value=left(document.form1.wk_exe.value,len(document.form1.wk_exe.value)-1)
         else
		document.form1.wk_exe.value=document.form1.wk_exe.value & "," & document.form1.exemen_w.value
         end if
	end if
  end if
	document.form1.exemen_w.value=""
end sub

sub wkradd()'���|�H��
  if document.form1.wkrmen_w.value="clear" then
   document.form1.all_worker.value=""
  else
	if document.form1.all_worker.value="" then
		document.form1.all_worker.value=document.form1.attmen_w.value
	else
         if instr(1,document.form1.all_worker.value,document.form1.wkrmen_w.value,1)>0 then
            document.form1.all_worker.value=replace(document.form1.all_worker.value,document.form1.wkrmen_w.value,"")
            document.form1.all_worker.value=replace(document.form1.all_worker.value,",,",",")
            if left(document.form1.all_worker.value,1)="," then document.form1.all_worker.value=right(document.form1.all_worker.value,len(document.form1.all_worker.value)-1)
            if right(document.form1.all_worker.value,1)="," then document.form1.all_worker.value=left(document.form1.all_worker.value,len(document.form1.all_worker.value)-1)
         else
		document.form1.all_worker.value=document.form1.all_worker.value & "," & document.form1.wkrmen_w.value
         end if
	end if
  end if
	document.form1.wkrmen_w.value=""
end sub	

sub menadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.men_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.men_w.value
	end if
end sub
sub dateadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.doing_date2.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.doing_date2.value
	end if
	'document.form1.doing_date1.value=document.form1.date_w.value
end sub
sub timeadd()
	'document.form1.time_w1.value=document.form1.time_w.value
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.time_w.value
	else	
		document.form1.wk_item.value=document.form1.time_w.value+" "+document.form1.wk_item.value
	end if
end sub
sub placeadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.place_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.place_w.value
	end if
end sub
sub thingadd()
	if document.form1.wk_item.value="" then
		document.form1.wk_item.value=document.form1.thing_w.value
	else	
		document.form1.wk_item.value=document.form1.wk_item.value+" "+document.form1.thing_w.value
	end if
end sub
sub item_chk()
	if document.form1.wk_item.value="" or document.form1.all_worker.value="" then
		ok=msgbox("�п�J�D���Ϊ��|�H���I�I",0,"���~ĵ�i")
	else
	end if
end sub
sub press_chk()

	str_err=""
	if document.form1.all_worker.value="" then str_err= str_err & chr(13) &"�п�J�i���|�H���j�I�I"  
	if document.form1.wk_exe.value="" then str_err= str_err & chr(13) &"�п�J�i����H���j�I�I"  
      pdate1=document.form1.doing_date1.value
      if pdate1="" or not(IsDate(pdate1) ) then
         p_chk1=0
      else
         p_chk1=1
      end if
	if p_chk1=0 then str_err= str_err & chr(13) &"�п�J�i�������j�I�I" 
	if document.form1.wk_item.value="" then str_err= str_err & chr(13) &"�п�J�i�D���j�I�I"  

	if not(trim(str_err)="") then
		ok=msgbox( str_err,0,"���~ĵ�i")	
	else
		hok=msgbox("�O�_�C�J�]���O�T���H�H",4+32+256,"�]���O�T��")
		if hok=6 then
			document.form1.headline1.value=10
		end if
		form1.submit
	end if
end sub

sub deadline_s()
   p_send_date=document.form_ipt.p_send_date.value
   p_dl_i=document.form_ipt.p_deadline.value
   p_dead_date=DateAdd("m", p_dl_i, p_send_date)
   document.form_ipt.p_dead_date.value=p_dead_date
end sub
sub month03_onClick()
   p_send_date=document.form_ipt.p_send_date.value
   p_dead_date=DateAdd("m", 3, p_send_date)
   document.form_ipt.p_dead_date.value=p_dead_date
end sub
sub month06_onClick()
   p_send_date=document.form_ipt.p_send_date.value
   p_dead_date=DateAdd("m", 6, p_send_date)
   document.form_ipt.p_dead_date.value=p_dead_date
end sub


sub pjnsel()
      if document.form1.wk_pjn.value="0" then
         document.form1.wk_group.value="�@��u�@"
         document.form1.wk_class.value=""
      else
         document.form1.wk_group.value="�M�פu�@"
         document.form1.wk_class.value=""
      end if
end sub

</script>
<script language=vbscript>
sub addaw_all()
	str_old="<%=worker%>"
	str_all="<%=stra_gp0%>"
    document.form1.all_worker.value=str_all
end sub
sub addaw_none()
	str_old="<%=worker%>"
    document.form1.all_worker.value=str_old
end sub
sub addaw_gp1()
'���`��{�H��
	str_gp1="<%=stra_gp1%>"
	pw1=document.form1.all_worker.value
	if instr(1,str_gp1,pw1,1)=0 then
	 	document.form1.all_worker.value=document.form1.all_worker.value & "," &str_gp1
	else
	 	document.form1.all_worker.value=str_gp1
	 end if
end sub
sub addaw_gp5()
'���~�H��
	str_gp5="<%=stra_gp5%>"
	 document.form1.all_worker.value=document.form1.all_worker.value & "," &str_gp5
end sub
sub addaw_gp6()
'�~�ȤH��
	str_gp6="<%=stra_gp6%>"
 	 document.form1.all_worker.value=document.form1.all_worker.value & "," &str_gp6
end sub
sub addaw_gp7()
'����|
	str_gp7="<%=stra_gp7%>"
	 document.form1.all_worker.value=document.form1.all_worker.value & "," &str_gp7
end sub
sub addsame_exe()
	str_old="<%=worker%>"
	'����H��
	str_all="<%=stra_gp0%>"
	'���~�H��
	str_gp5="<%=stra_gp5%>"
	'�~�ȤH��
	str_gp6="<%=stra_gp6%>"
'�P����H�� ps_exe
ps_exe=document.form1.wk_exe.value
if ps_exe="" then
	hok=msgbox("�п�J����H���I�I",0+48+0,"���~�T��")
else
	if instr(1,ps_exe,"����H��",1) >0 then
		ps_allworker=str_all	
	elseif instr(1,ps_exe,"�~�ȥ���",1) >0 then
		ps_allworker1=str_gp6	
		ps_allworker2=replace(ps_exe,"�~�ȥ���","")
		ps_allworker2=replace(ps_allworker2,",,",",")
		if left(ps_allworker2,1)="," then ps_allworker2=right(ps_allworker2,len(ps_allworker2)-1) 
		if right(ps_allworker2,1)="," then ps_allworker2=left(ps_allworker2,len(ps_allworker2)-1) 
		if ps_allworker2="" then
			ps_allworker=ps_allworker1
		else
			ps_allworker=ps_allworker1 & "," &	ps_allworker2
		end if
	elseif instr(1,ps_exe,"���~����",1) >0 then
		ps_allworker1=str_gp5	
		ps_allworker2=replace(ps_exe,"���~����","")
		ps_allworker2=replace(ps_allworker2,",,",",")
		if left(ps_allworker2,1)="," then ps_allworker2=right(ps_allworker2,len(ps_allworker2)-1) 
		if right(ps_allworker2,1)="," then ps_allworker2=left(ps_allworker2,len(ps_allworker2)-1) 
		if ps_allworker2="" then
			ps_allworker=ps_allworker1
		else
			ps_allworker=ps_allworker1 & "," &	ps_allworker2
		end if	
	else
		ps_allworker=ps_exe		
	end if
end if
	if instr(1,ps_allworker,str_old,1)=0 then 
		if ps_allworker="" then
			ps_allworker=str_old 
		else
			ps_allworker=str_old & "," & ps_allworker
		end if 
	end if
document.form1.all_worker.value=ps_allworker
end sub

sub attadddp()    '==20170606�s�W===�����ﶵ===�X�u�H��=====
  if document.form1.attmen_dp.value="clear" then
   document.form1.wk_att.value=""
  else
	if document.form1.wk_att.value="" then
		document.form1.wk_att.value=document.form1.attmen_dp.value
	else
         if instr(1,document.form1.wk_att.value,document.form1.attmen_dp.value,1)>0 then
            document.form1.wk_att.value=replace(document.form1.wk_att.value,document.form1.attmen_dp.value,"")
            document.form1.wk_att.value=replace(document.form1.wk_att.value,",,",",")
            if left(document.form1.wk_att.value,1)="," then document.form1.wk_att.value=right(document.form1.wk_att.value,len(document.form1.wk_att.value)-1)
            if right(document.form1.wk_att.value,1)="," then document.form1.wk_att.value=left(document.form1.wk_att.value,len(document.form1.wk_att.value)-1)
         else
		document.form1.wk_att.value=document.form1.wk_att.value & "," & document.form1.attmen_dp.value
         end if
	end if
  end if
	document.form1.attmen_dp.value=""
end sub

sub exeadddp()    '==20151209�s�W===�����ﶵ===����H��=====
  if document.form1.exemen_dp.value="clear" then
   document.form1.wk_exe.value=""
  else
	if document.form1.wk_exe.value="" then
		document.form1.wk_exe.value=document.form1.exemen_dp.value
	else
         if instr(1,document.form1.wk_exe.value,document.form1.exemen_dp.value,1)>0 then
            document.form1.wk_exe.value=replace(document.form1.wk_exe.value,document.form1.exemen_dp.value,"")
            document.form1.wk_exe.value=replace(document.form1.wk_exe.value,",,",",")
            if left(document.form1.wk_exe.value,1)="," then document.form1.wk_exe.value=right(document.form1.wk_exe.value,len(document.form1.wk_exe.value)-1)
            if right(document.form1.wk_exe.value,1)="," then document.form1.wk_exe.value=left(document.form1.wk_exe.value,len(document.form1.wk_exe.value)-1)
         else
		document.form1.wk_exe.value=document.form1.wk_exe.value & "," & document.form1.exemen_dp.value
         end if
	end if
  end if
	document.form1.exemen_dp.value=""
end sub

sub wkradddp() '==20151209�s�W===�����ﶵ===���|�H��=====
	str_old="<%=worker%>"
  if document.form1.wkrmen_dp.value="clear" then
   document.form1.all_worker.value=str_old
  else
	if document.form1.all_worker.value="" then
		document.form1.all_worker.value=document.form1.wkrmen_dp.value
	else
         if instr(1,document.form1.all_worker.value,document.form1.wkrmen_dp.value,1)>0 then
            document.form1.all_worker.value=replace(document.form1.all_worker.value,document.form1.wkrmen_dp.value,"")
            document.form1.all_worker.value=replace(document.form1.all_worker.value,",,",",")
            if left(document.form1.all_worker.value,1)="," then document.form1.all_worker.value=right(document.form1.all_worker.value,len(document.form1.all_worker.value)-1)
            if right(document.form1.all_worker.value,1)="," then document.form1.all_worker.value=left(document.form1.all_worker.value,len(document.form1.all_worker.value)-1)
         else
		document.form1.all_worker.value=document.form1.all_worker.value & "," & document.form1.wkrmen_dp.value
         end if
	end if
  end if
	document.form1.wkrmen_dp.value=""
end sub
sub addexe_none()
    document.form1.wk_exe.value=str_old
end sub

sub addatt_none()
    document.form1.wk_att.value=str_old
end sub

sub addsameatt_exe()
	str_old="<%=worker%>"
	'����H��
	str_all="<%=stra_gp0%>"
	'���~�H��
	str_gp5="<%=stra_gp5%>"
	'�~�ȤH��
	str_gp6="<%=stra_gp6%>"
'�P����H�� ps_exe
ps_exe=document.form1.wk_exe.value
if ps_exe="" then
	hok=msgbox("�п�J����H���I�I",0+48+0,"���~�T��")
else
	if instr(1,ps_exe,"����H��",1) >0 then
		ps_allworker=str_all	
	elseif instr(1,ps_exe,"�~�ȥ���",1) >0 then
		ps_allworker1=str_gp6	
		ps_allworker2=replace(ps_exe,"�~�ȥ���","")
		ps_allworker2=replace(ps_allworker2,",,",",")
		if left(ps_allworker2,1)="," then ps_allworker2=right(ps_allworker2,len(ps_allworker2)-1) 
		if right(ps_allworker2,1)="," then ps_allworker2=left(ps_allworker2,len(ps_allworker2)-1) 
		if ps_allworker2="" then
			ps_allworker=ps_allworker1
		else
			ps_allworker=ps_allworker1 & "," &	ps_allworker2
		end if
	elseif instr(1,ps_exe,"���~����",1) >0 then
		ps_allworker1=str_gp5	
		ps_allworker2=replace(ps_exe,"���~����","")
		ps_allworker2=replace(ps_allworker2,",,",",")
		if left(ps_allworker2,1)="," then ps_allworker2=right(ps_allworker2,len(ps_allworker2)-1) 
		if right(ps_allworker2,1)="," then ps_allworker2=left(ps_allworker2,len(ps_allworker2)-1) 
		if ps_allworker2="" then
			ps_allworker=ps_allworker1
		else
			ps_allworker=ps_allworker1 & "," &	ps_allworker2
		end if	
	else
		ps_allworker=ps_exe		
	end if
end if
	'if instr(1,ps_allworker,str_old,1)=0 then 
		'if ps_allworker="" then
			'ps_allworker=str_old 
		'else
			'ps_allworker=str_old & "," & ps_allworker
		'end if 
	'end if
document.form1.wk_att.value=ps_allworker
end sub

</script>

</center>
</body>
</html>
