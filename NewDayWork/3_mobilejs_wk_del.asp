<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<!-- #Include file = "./include/array_wkclass.inc" -->
<!-- #Include file = "./include/array_wkgroup.inc" -->
<!-- #Include file = "./include/workinput.inc" -->
<!-- #Include file = "./misc_data/array_place.inc" -->	
<!-- #Include file = "./misc_data/array_thing.inc" -->	
<!-- Include file = "./misc_data/array_writer.inc" -->	
<!-- #Include file = "./include/array_pjn.inc" -->
<%
'�DIE�s�����Τ�������s��\��
	'Ū���H���m�W
	worker = Session("worker")
%>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
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
wk_doer=rstObj1.fields("wk_doer")         '���|�H��
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_class1=rstObj1.fields("wk_class")
wk_group1=rstObj1.fields("wk_group")
wk_exe1=rstObj1.fields("wk_exe")
wk_pjid=rstObj1.fields("pj_id")          '�M�צW��id
wk_pjn=rstObj1.fields("pj_02")          '�M�צW��

if wk_group1="�M�פu�@" and doing_date1 < date() then doing_date1=date()

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
		font-family:'�з���';		/*�r��*/
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
		font-family: '�з���';
		font-size: 3.3mm;
		line-height: 18px;
		color:blue;
		font-weight:bold;
		}
TD.myd{
		font-family: '�з���';
		font-size: 3.3mm;
		line-height: 18px;
		background-color:#f0ffff;
		}     
    
-->
</style>
<script type="text/javascript">
function validateForm()
{
var x1=document.forms["form1"]["wk_item"].value;
var x2=document.forms["form1"]["wk_content"].value;
var x3=document.forms["form1"]["doing_date1"].value;
var x4=document.forms["form1"]["all_worker"].value;
kk1=notEmpty(document.getElementById('t_item'), '�п�J�D���C');
kk2=notEmpty(document.getElementById('t_content'), '�п�J���e�C');
kk3=notEmpty(document.getElementById('t_allworker'), '�п�J���|�H���C');
kk4=isDates(document.getElementById('do_date'), '�Х��T��J����榡2011/01/01�C');
//if (x1==null || x1=="" || x2==null || x2=="" || x3==null || x3=="" || x4==null || x4=="" )
if (kk1 || kk2 || kk3 || kk4 )
  {
//  alert("�нT���J��ơI�I�D���B���e�B�u�@����B���|�H��" );
  return false;
  }
    {
  alert("��Ƥw�T���J�I�I"  );
  return false;
  }
}
function notEmpty(elem, helperMsg){
	if(elem.value.length == 0){
		alert(helperMsg);
		elem.focus();
		return false;
	}
	return true;
}
// If the element's string matches the regular expression it is all numbers
function isNumerics(elem, helperMsg){
	var numericExpression = /^[0-9]+$/;
	if(elem.value.match(numericExpression)){
		return true;
	}else{
		alert(helperMsg);
		elem.focus();
		return false;
	}
}
// If the element's string matches the regular expression it is all numbers    /\d{4}\/\d{2}-\/\d{2}/
function isDates(elem, helperMsg){
	var dateExpression = /\d{4}\/[01]\d{1}\/[0123]\d{1}/;
	var dlen=elem.length
	if(elem.value.match(dateExpression)){
		if (dlen == 10){
		return true;
		}else{
      		alert(helperMsg);
      		elem.focus();
      		return false;		
		}
	}else{
		alert(helperMsg);
		elem.focus();
		return false;
	}
}
</script>

</HEAD>
<BODY>
<center>
<font color=red>
</font>
<form name="form1" action="3_mobilejs_wk_del_ok.asp" method="post" >
<input type="hidden" name="worker1" value="<%=worker%>" >
<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<table border=1 cellspacing=0 cellpadding=0>
<col width=100>
<col width=130>
<col width=100>
<col width=130>
<col width=100>
<col width=130>
<tr>
	<td colspan=6 align="center">
         �T�{�R�����i��ơH	
         <input type="submit" name="submit" value="�T�w�R��" >
		<input type="button" name="bkpg" value="�^�W�@��" style="cursor:hand;" onclick="parent.location.href='javascript:history.back()'" >
	</td>
<tr>
</table>
<%
function showspace(ztxt)
   if ztxt="" or isnull(ztxt) then
      pztxt="&nbsp;"
   else
      pztxt=ztxt
   end if
   showspace=pztxt
end function
%>
<table border=1 cellspacing=0 cellpadding=0>
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<col width=120>
<col width=120 style="padding-left:5px;">
<tr>
	<td align="center" colspan=2 rowspan=2><font size=4 color="red"><b>��ܳ�@�u�@��</b></font></td>
	<td align="right">�u�@�s�աG</td>
	<td><%=showspace(wk_group1)%>
	<!-- <input type='text' name='wk_group' value='<%=wk_group%>' style="width:100%;" readonly> -->
	</td>
	<td align="right">�M�צW�١G</td>
	<td><%=showspace(wk_pjn)%>
<!-- 	<input type='text' name='wk_pjn' value='<%=wk_pjn%>' style="width:100%;" readonly> -->
	</td>
</tr>

<tr>
<!-- 	<td align="center" colspan=2><font size=4 color="red"><b><%=wk_group%></font></td> -->
	<td align="right">�u�@�s���G</td>
	<td><%=showspace(wk_id)%>
	<!-- <input type='text' name='wk_id' value='<%=wk_id%>' style="width:100%;" readonly> -->
	</td>
	<td align="right">�u�@�����G</td>
	<td><%=showspace(wk_class)%>
	<!-- <input type='text' name='wk_class' value='<%=wk_class%>' style="width:100%;" readonly> -->
	</td>
</tr>

<tr>
	<td align="right">���i�̡G</td>
	<td><%=showspace(wk_order)%>
	<!-- <input type='text' name='wk_order1' value='<%=wk_order%>' style="width:100%;" readonly> -->
	</td>
	<td align="right">���i����G</td>
	<td><%=showspace(undo_date1)%>
	<!-- <input type='text' name='undo_date1' value='<%=undo_date1%>' style="width:100%;" readonly> -->
	</td>
	<td align="right">�������G</td>
	<td><%=showspace(doing_date1)%>
<!-- 	<input type='text' name='doing_date1' value='<%=doing_date1%>' style="width:100%;" readonly> -->
	</td>
</tr>
<tr>
	<td align="right">
	���|�H���G
	</td>
	<td colspan=5><%=showspace(wk_doer)%>
<!-- 	<input type='text' name='wk_doer' value='<%=wk_doer%>' style="width:100%;" readonly> -->
	</td>
</tr>
<tr>
	<td align="right">
	�����H���G
	</td>
	<td colspan=5><%=showspace(wk_checker)%>
<!-- 	<input type='text' name='wk_checker' value='<%=wk_checker%>' style="width:100%;" readonly> -->
	</td>
</tr>
<tr>
	<td align="right">
	�������H���G
	</td>
	<td colspan=5><%=showspace(wk_undoer)%>
<!-- 	<input type='text' name='wk_undoer' value='<%=wk_undoer%>' style="width:100%;" readonly> -->
	</td>
</tr>
<tr>
	<td align="right">
	�D���G
	</td>
	<td colspan=5><%=showspace(wk_item)%>
<!-- 	<input type='text' name='wk_item' value='<%=wk_item%>' style="width:100%;" readonly> -->
	</td>
</tr>
<tr>
	<td align="right" valign="top">
	���椺�e�G
	</td>
	<td colspan=5>
	<%
	if wk_content="" or isnull(wk_content) then
	  wk_content_a=wk_content
	else
	  wk_content_a=replace(wk_content,chr(13),"<br>")
	end if
	response.write  wk_content_a
	%>
<!-- 	<TEXTAREA name="wk_content" rows="10" style="width:100%;" readonly><%=wk_content%></TEXTAREA>
 -->
 	</td>
</tr>



</table>

</form>

</center>
</body>
</html>
