<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
	wk_chk=Request("wk_chk")
	strbackURL=session("strbackURL")
%>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
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
wk_exe=rstObj1.fields("wk_exe")           '����H��
wk_att=rstObj1.fields("wk_att")           '�X�u�H��
wk_pjn=rstObj1.fields("pj_02")   '�M�צW��
pwk_password=rstObj1.fields("wk_password")   '�[�K��r
wk_headline=rstObj1.fields("headline")'�]���O

%>
<%
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
' <%
' '�P�_�O�_�OIE�Τ��
' dim u,b
' set u=Request.ServerVariables("HTTP_USER_AGENT")
' 'response.write u
' 'response.write "<hr>"
' 'response.end
' '
' ck_MSIE=instr(1,u,"MSIE",1)
' ck_IE=instr(1,u,"IE",1)

' ck_Chrome=instr(1,u,"Chrome",1)
' ck_Firefox=instr(1,u,"Firefox",1)

' ck_Safari=instr(1,u,"Safari",1)
' ck_Firefox=instr(1,u,"Firefox",1)

' if ck_MSIE+ck_IE>0 then
'    'IE�s����
'    ck_mobile=0
' elseif ck_Chrome >0 then
'    'Chrome�s����
'    ck_mobile=1
' elseif ck_Firefox>0 then
'    'Firefox�s����
'    ck_mobile=1
' elseif ck_Safari>0 then
'    'Safari�s����
'    ck_mobile=1
' else
'    ck_mobile=1
' end if

' if ck_mobile=1 then
'       nexturl="3_mobilejs_wk_show.asp?wk_id="& wk_id&"&wk_chk="&wk_chk
'       response.redirect(nexturl)
' else
' end if

' 'set b=new RegExp
' 'b.Pattern="firefox|chrome|safari|mobile"
' 'b.Pattern="safari|mobile"
' 'b.IgnoreCase=true
' 'b.Global=true
' 'Set matchesb = b.Execute(u)
' 'if b.test(u) then               '�DIE�s����
' '      response.redirect("http://detectmobilebrowser.com/mobile")
' '      response.write "b="& matchesb(0).value &"<hr>"
' '      response.write "b.test(u)="&b.test(u)&"<hr>"
' '      response.write "�s�����G"& matchesb(0).value & "<hr>"
'       '�DIE
'       'nexturl="3_mobilejs_wk_show.asp?wk_id="& wk_id&"&wk_chk="&wk_chk
'       'response.redirect(nexturl)
' 'else
' '      response.write "b.test(u)="&b.test(u)&"<hr>"
' '      response.write "�s�����G"&"IE<hr>"
' 'end if
' 'response.end
' %>
<%
'�O�_�[�K���
if wk_chk="ok" or pwk_password="" or isnull(pwk_password) then
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
<BODY>
<center>
<form name="form1" action="" method="post">
<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<input type="hidden" name="worker1" value="<%=worker%>">
<input type="hidden" name="wk_ordera" value="<%=wk_order%>">
<%if wk_group="�@��u�@" then%>
<!-- Include file = "./include/toolbar_show.inc" -->
<script language=vbscript>
'sub edit_click()
'	wk_id=document.form1.wk_id1.value
'	location.href="./wk_edit.asp?wk_id="&wk_id
'end sub
sub delete_click()
	worker=document.form1.worker1.value
	wk_order=document.form1.wk_ordera.value
	if worker=wk_order then
		ok=msgbox("�O�_�T�w�n�R����ơH",1,"�R��ĵ�i")
		if ok=1 then
			wk_id=document.form1.wk_id1.value
			location.href="./wk_del_ok.asp?wk_id="&wk_id
		end if
	else
		ok=msgbox("�A���O���u�̡A�L�k�R�������u�@�I�I",0,"���~ĵ�i")
	end if
end sub
sub done_click()
	ok=msgbox("�O�_�T�w�n�����u�@�H",1,"�T�{ĵ�i")
	if ok=1 then
		wk_id=document.form1.wk_id1.value
		location.href="./wk_done_ok.asp?wk_id="&wk_id
	end if
end sub
sub readd_click()
	ok=msgbox("�O�_�T�w�n���s���i�u�@�H",1,"�T�{ĵ�i")
	if ok=1 then 
		wk_id=document.form1.wk_id1.value
		location.href="./wk_readd.asp?wk_id="&wk_id
	end if
end sub
sub gpchange_click()
	ok=msgbox("�O�_�T�w�n�ର�M�פu�@�H",1,"�T�{ĵ�i")
	if ok=1 then
		wk_id=document.form1.wk_id1.value
		location.href="./wk_gpchg_special.asp?wk_id="&wk_id
	end if
end sub
</script>

<script type="text/javascript">
function edit_click()
{
   var x1=document.forms["form1"]["wk_id1"].value;
   str_url="./wk_edit.asp?wk_id="+x1;
   // alert(str_url);
   location.href = str_url ;
}

</script>

<center>

<table border=0 cellspacing=0 cellpadding=0>
<col span=8 style="width:90px;text-align:center;">
<tr width=720>	
	<td>
		<input type=button name="bkpg" value="�^�W�@��" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="parent.location.href='javascript:history.back()'" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="edit" value="�s�פu�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="edit_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="delete" value="�R���u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="delete_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
<% if wk_class="Z" then %>
		<input type=button name="done" value="�����u�@" disabled style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
<% else %>
		<input type=button name="done" value="�����u�@" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="done_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
<% end if %>
	</td>
	<td>
		<input type=button name="readd" value="���s���i" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="readd_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="wkprint" value="�C�L���e" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="window.open('./wkprint_si.asp?wk_id=<%=wk_id%>')"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="gpchange" value="�ର�M��" style="cursor:hand;background-color:'#d3d3d3';color:blue;width:100%;" onclick="gpchange_click()" onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#d3d3d3';">
	</td>
	<td>
		<input type=button name="wkattfile" value="�W�Ǫ���" style="cursor:hand;background-color:'#77FFEE';color:blue;width:100%;" onclick="location.href='./1_ulf_form.asp?wk_id=<%=wk_id%>'"  onmouseover="javascript:this.style.background='#FFd700';" onmouseout="javascript:this.style.background='#77FFEE';">
	</td>
</tr>
</table>
</center>
<%else%>
<!-- #Include file = "./include/toolbar_pj_show.inc" -->
<%end if%>
<!-- #Include file = "./include/wk_show_form.inc" -->

<%
'���[�ɮצC��
' �s��Access��Ʈwdaywork.mdb
DBpath=Server.MapPath("./database/attach_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="file_data"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id &" and del_ok = false order by fl_date desc"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
else
%>
<table border=1 cellspacing=0 cellpadding=0 width=750 bgcolor="#CCEEFF">
<col width=40 style="text-align:center;">
<col width=280 style="padding-left:5px;text-align:left;">
<col width=210 style="padding-left:5px;text-align:left;">
<col width=90 style="text-align:center;">
<col width=90 style="text-align:center;">
<tr>
<td colspan=5>����C��</td>
</tr>
<tr>
<td >�Ǹ�</td>
<td align=center >�ɮ׻���</td>
<td align=center >�ɮצW��  [�W�Ǫ�]</td>
<td >���ɤ��</td>
<td >�\��</td>
</tr>
<%
	'�C�X��ƶ���
	rstobj1.MoveFirst
	for fi=1 to totalput
	'Ū�����
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")
		pfl_item=rstObj1.fields("fl_item")
		pfl_inputer=rstObj1.fields("fl_inputer")
		pfl_history= rstObj1.fields("fl_history")
		pfl_date=rstObj1.fields("fl_date")
		str_none=pwk_id&"_"
		str_pfl_name=right(pfl_name,len(pfl_name)-len(pwk_id)-1)
%>
<tr>
<td ><%=fi%></td>
<td >
<a href="./1_ulf_item_edit.asp?fl_id=<%=pfl_id%>" target="_self" title="�ק��ɮ׻����C" ><img src="./img/change.png" style="vertical-align:middle;height:16px;cursor:hand;border:0;" ></a>
<%=pfl_item%>
</td>
<td ><a href="./file_att/<%=pfl_name%>" target="_blank" title="<%=pfl_history%>"><%=str_pfl_name%></a>  [<%=pfl_inputer%>]</td>
<td ><%=pfl_date%></td>
<td >
<input type="button" name="addfile" value="�s"  onclick="file_add('<%=pwk_id%>')" title="�u�@���� [ wk_id=<%=pwk_id%> ] �s�W�ɮסC">
<input type="button" name="delfile" value="�R"  onclick="file_del('<%=pfl_id%>')" title="�N�ɮקR���C">
<!-- <a href="1_ulf_form.asp?wk_id=<%=pwk_id%>" title="�s�W�ɮשΧ�s�ɮסC">�s</a>
<a href="1_ulf_del.asp?fl_id=<%=pfl_id%>" title="�R���ɮסC">�R</a> -->
</td>
</tr>
<%
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
</form>
<hr>
<!-- �s������ơG<%=u%> -->
</center>
<script language=vbscript>
sub file_add(ppwk_id)
	ok=msgbox("�O�_�T�w�n�s�W��ơH"&chr(13)&"1_ulf_form.asp?wk_id="&ppwk_id,1,"�s�Wĵ�i")
	if ok=1 then 
		location.href="1_ulf_form.asp?wk_id="&ppwk_id
	end if
end sub
sub file_del(ppfl_id)
	ok=msgbox("�O�_�T�w�n�R����ơH"&chr(13)&"1_ulf_del.asp?fl_id="&ppfl_id,1,"�R��ĵ�i")
	if ok=1 then 
		location.href="1_ulf_del.asp?fl_id="&ppfl_id
	end if
end sub

</script>
</body>
</html>
<%
else     '�[�K����J�[�K��r
%>
<html>
<head>
<title>�K�X�ˬd</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�s�ө���';background-color :'#FFFEEE'}
input{
	font-family:'�s�ө���';
	font-size:12pt;
	}
select{font-family:'�s�ө���';font-size:10pt;cursor:hand;}
.itxt{
	font-family:'�s�ө���';
	font-size:12pt;
	width:100%;
	height:100%;
	}
input.imenu { 
	/*font-size:15px;				/*�r��j�p*/
	/*font-weight:bold;
	cursor:hand;				/*��ЧΦ�*/ 
	background-color:'<%=botton_color%>'; 		
	margin:0 0 0 0;		/*��t�W�U���k*/
	width:100px;
	/*height:100%;*/
	color:#000000;
	letter-spacing:2px;
	cursor:hand;
     }
td{
	margin:0 0 0 0;		/*��t�W�U���k*/
	border-color:'#808080'; /*���~���C��*/ 
	border-style:solid;		/*���~�ؽu��*/
	border-width:1px;		/*���~�ثp��*/  
	vertical-align:middle;	/*�r�髫������覡*/
	/*font-size:15px;*/ 
	}
table{	
	margin:0 0 0 0;		/*��t�W�U���k*/
	border-collapse:collapse; 	/*��اΦ����X*/
	}
input.itext { 
	font-size:3.5mm;				/*�r��j�p*/
	/*cursor:hand;				/*��ЧΦ�*/ 
	width:100%;
	height:5mm;
	background-color:'#ffeedd'; 		/*�~���C��*/
	margin:0 0 0 0;		/*��t�W�U���k*/
	color:black;
	text-align:right;
     }

--></style>
</head>
<body>
<center>
<form name="form_login" method=post action="">
<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<input type="hidden" name="wk_pwd" value="<%=pwk_password%>">
<table border=0 cellspacing=0 cellpadding=0>
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<tr>
	<td align="center" colspan=2 ><font size=4 color="red"><b>��ܳ�@�u�@��</b></font></td>
	<td align="right">�u�@�s�աG</td>
	<td><!-- <%=showspace(wk_group)%> -->
	<input type='text' name='wk_group' value='<%=wk_group%>' style="width:100%;" readonly>
	</td>
	<td align="right">�M�צW�١G</td>
	<td><!-- <%=showspace(wk_pjn)%> -->
 	<input type='text' name='wk_pjn' value='<%=wk_pjn%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">���i�̡G</td>
	<td><!-- <%=showspace(wk_order)%>-->
	<input type='text' name='wk_order1' value='<%=wk_order%>' style="width:100%;" readonly>
	</td>
	<td align="right">���i����G</td>
	<td><!-- <%=showspace(undo_date1)%> -->
	<input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly>
	</td>
	<td align="right">�������G</td>
	<td><!--<%=showspace(doing_date1)%>-->
 	<input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	����H���G
	</td>
	<td colspan=5><!--<%=showspace(wk_checker)%> -->
 	<input type='text' name='wk_exe' value='<%=wk_exe%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right" style="background-color:#FFBFFF;">
	�X�u�H���G
	</td>
	<td colspan=5>
 	<input type='text' name='wk_att' value='<%=wk_att%>' style="width:100%;" readonly  onkeydown="javascript:if(window.event.keyCode==8) return false;">
	</td>
</tr>
<tr>
	<td align="right">
	���|�H���G
	</td>
	<td colspan=5><!--<%=showspace(wk_doer)%> -->
 	<input type='text' name='wk_doer' value='<%=wk_doer%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	�����H���G
	</td>
	<td colspan=5><!--<%=showspace(wk_checker)%> -->
 	<input type='text' name='wk_checker' value='<%=wk_checker%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	�������H���G
	</td>
	<td colspan=5><!--<%=showspace(wk_undoer)%> -->
 	<input type='text' name='wk_undoer' value='<%=wk_undoer%>' style="width:100%;" readonly>
	</td>
</tr>
<tr>
	<td align="right">
	�D���G
	</td>
	<td colspan=5><!--<%=showspace(wk_item)%>-->
 	<input type='text' name='wk_item' value='<%=wk_item%>' style="width:100%;" readonly>
	</td>
</tr>
</table>
<hr color=red>
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">�i�п�J�[�K��r�H�˵��������e�j</font>
<table border=0 cellspacing=0 cellpadding=2 style="width:300px" >
<col style="width:100px;font-size:4mm;" align=center>
<col style="width:200px;font-size:4mm;" align=center>
<tr>
<td>�[�K��r�G</td>
<td><input type="text" style="text-align:left;" name="wkr_pwd" value="" maxlength="10" ></td>
</tr>
<tr>
<td colspan=2>
	<input type="button" name="submit01" value="�˵��������e" onclick="check_password()">
	<input type="button" name="reset01" value="�^�W��" onclick="back_url()" >
</td>
</tr>
</table>
<hr color=red>
</body>
</html>
<script language='Vbscript'>
<!--
sub check_password()
   chk_str=document.form_login.wk_pwd.value
   ipt_str=document.form_login.wkr_pwd.value
   pwk_id=document.form_login.wk_id1.value
   if chk_str=ipt_str then
      str_url="./wk_show.asp?wk_id="&pwk_id&"&wk_chk=ok"
      'MyVar = MsgBox ("wk_id="&pwk_id&"�Cchk_str="&chk_str, 16, "���~�T��")
      location.href=str_url
   else
      str_url="<%=strbackURL%>"
      MyVar = MsgBox ("�[�K��r���~�I�I"&chr(13)&"�^��W�@���I�I", 16, "���~�T��")
      location.href=str_url
   end if

end sub
sub back_url()
   back_str="<%=strbackURL%>"
   location.href=back_str
end sub
-->
</script>
<%
end if
%>