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
stra_gp1="����,���,���,����,����,���z,�f�S,���F,�@�q,�_��"   '������{�M��

'�DIE�s�����Τ�������s�W�\��
	'Ū���H���m�W
	worker = Session("worker")
	if worker="" or isnull(worker) then response.redirect("./firstpage.asp")
	datecode=request("datecode")
	if datecode="" then datecode=date()
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
str_allworker=""
for i=1 to worker_no
   if str_allworker="" then
      str_allworker=worker_a(i-1)
   else
      str_allworker=str_allworker & "," & worker_a(i-1)
   end if
next
%>
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_id=Request("wk_id")
%>
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
p_undo_date1=rstObj1.fields("undo_date1")
p_doing_date1=rstObj1.fields("doing_date1")
p_done_date1=rstObj1.fields("done_date1")
p_wk_item=rstObj1.fields("wk_item")
p_wk_content=rstObj1.fields("wk_content")
p_wk_order=rstObj1.fields("wk_order")
p_wk_doer=rstObj1.fields("wk_doer")
p_wk_checker=rstObj1.fields("wk_checker")
p_wk_undoer=rstObj1.fields("wk_undoer")
p_wk_class=rstObj1.fields("wk_class")
p_wk_group=rstObj1.fields("wk_group")
p_wk_exe=rstObj1.fields("wk_exe")
p_wk_pjn=rstObj1.fields("pj_02")   '�M�צW��
p_wk_att=rstObj1.fields("wk_att")           '�X�u�H��

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
<style type="text/css">
<!--
input {
font-size:5mm;
}
a {
font-size:5mm;
}
tr{
/*height:120px;*/
}
td{
font-size:5mm;
text-align:center;
}
TEXTAREA{
font-size:5mm;
}
SELECT{
font-size:5mm;
}
input.checkbox {
font-size:5mm;
}

-->
</style>


</HEAD>
<BODY >
<center>
<form name="form1" action="3_mobilejs_wk_edit_ok.asp" method="post" >
	<input type="hidden" name="wk_id1" value="<%=wk_id%>">
<font style="font-size:5mm;" color="red"><b><%=worker%>�u�@���i��ק�</b></font>
<table border=1 cellspacing=0 cellpadding=0 style="width:600px;">
<col style="width:100px;">
<col style="">
<tr>
	<td align="right"><font color="red">������</font></td>
	<td ><input type='text' name="doing_date1" id="do_date" value="<%=p_doing_date1%>" style="width:100%;"></td>
</tr>
<tr>
	<td align="right">	<font color="red">�D��</font></td>
	<td >	<input type='text' name='wk_item' id="t_item" value='<%=p_wk_item%>' style="width:100%;"  >	</td>
</tr>
<tr>
	<td align="right">
	<font color="red">����H��</font>
	</td>
	<td colspan=5 style="text-align:left;">
	<input type='text' name='wk_exe' id='t_exe' value='<%=p_wk_exe%>' style="width:50%;">
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
		<br>(�п�J����ѻP�H��)
	</td>
</tr>
<tr>
	<td align="right">	<font color="red">���椺�e</font></td>
	<td ><TEXTAREA name="wk_content" id="t_content" rows="5" style="width:100%;" ><%=p_wk_content%></TEXTAREA></td>
</tr>
<tr>
	<td align="right">	<font color="red">���|�H��</font>
      	<input type="button" name="press" value="������{" onclick="showeagle()">
   </td>
	<td > <font style="font-size:5mm;">
<%
	for i=1 to worker_no
'	  if worker=worker_a(i-1) then 'p_wk_doer
	  if instr(1,p_wk_doer,worker_a(i-1),1)>0 then 
	     str_chk="checked"
	  else
	     str_chk=""
	  end if
%>
<input type="checkbox" name="all_worker" id="t_allworker" value="<%=worker_a(i-1)%>" <%=str_chk%> ><%=worker_a(i-1)%>
 <%
      if ( i mod 7)=0 then response.Write "<br>"
	next
%>	</font>
	</td>
</tr>

<tr>
	<td colspan=2 align="center">
	<input type="button" name="press" value="�T�w���i" onclick="validateForm()">
	<input type="reset" name="cancel" value="�M�����" >
	</td>
<tr>
</table>
<font color=red>
</font>
</form>
<script type="text/javascript">
function validateForm(){
//alert("�нT���J��ơI�I�D���B���e�B�u�@����B���|�H���B����H��" );
var x1=document.forms["form1"]["wk_item"].value;
//var x2=document.forms["form1"]["wk_content"].value;
var x3=document.forms["form1"]["doing_date1"].value;
var x4=document.forms["form1"]["all_worker"];
var x5=document.forms["form1"]["wk_exe"].value;
//alert("x1="+x1+"�Cx2="+x2+"�Cx3="+x3+"�Cx4="+x4+"�Cx5="+x5  );
kk1=notEmpty(document.getElementById('t_item'), '�п�J�D���C');
//kk2=notEmpty(document.getElementById('t_content'), '�п�J���e�C');
kk3=isDatePart(document.getElementById('do_date'), '�Х��T��J����榡2011/01/01�C');
//kk3=isDates(document.getElementById('do_date'), '�Х��T��J����榡2011/01/01�C');
kk4=notEmpty(document.getElementById('t_allworker'), '�п�J���|�H���C');
kk5=notEmpty(document.getElementById('t_exe'), '�п�J����H���C');
//alert("kk1="+kk1+"�Ckk2="+kk2+"�Ckk3="+kk3+"�Ckk4="+kk4+"�Ckk5="+kk5  );
//if (kk1 && kk2 && kk3 && kk4 && kk5 )  {
if (kk1 && kk3 && kk4 && kk5 )  {
  alert("��Ƥw�T���J�I�I"  );
  document.forms["form1"].submit();
  //return (true);
  } else {
  alert("�нT���J��ơI�I�D���B�u�@����B���|�H���B����H��" );
  //return (true);
  }
}
function notEmpty(elem, helperMsg){
	if(elem.value.length == 0 || elem.value == ""){
		alert(helperMsg);
		//elem.focus();
		return false;
	} else{
	     return true;
	}
}
// If the element's string matches the regular expression it is all numbers
function isNumerics(elem, helperMsg){
	var numericExpression = /^[0-9]+$/;
	if(elem.value.match(numericExpression)){
		return true;
	}else{
		alert(helperMsg);
		//elem.focus();
		return false;
	}
}
//=============================================================
// * �P�_�@�Ӧr��O�_���X�k������榡�GYYYY-MM-DD
// */
function isDatePart(elem, helperMsg){
  var parts;
  var dateStr=elem.value  ;

  if(dateStr.indexOf("-") > -1){
    parts = dateStr.split('-');
  }else if(dateStr.indexOf("/") > -1){
    parts = dateStr.split('/');
  }else{
   alert(helperMsg);
    return false;
  }

  if(parts.length < 3){
  //������������\�ʤ֦~�B��B�餤������@��
    alert(helperMsg);
    return false;
  }

  for(i = 0 ;i < 3; i ++){
  //�p�G�c��������Y�ӳ������O�Ʀ�A�h��^false
    if(isNaN(parts[i])){
      alert(helperMsg);
      return false;
    }
  }

  y = parts[0];//�~
  m = parts[1];//��
  d = parts[2];//��

  if(y > 3000){
  alert(helperMsg);
    return false;
  }

  if(m < 1 || m > 12){
    alert(helperMsg);
    return false;
  }

  switch(d){
    case 29:
      if(m == 2){
      //�p�G�O2���
        if( (y / 100) * 100 == y && (y / 400) * 400 != y){
          //�p�G�~����Q100�㰣������Q400�㰣 (�Y�|�~)
        }else{
          alert(helperMsg);
          return false;
        }
      }
      break;
    case 30:
      if(m == 2){
      //2��S��30��
        alert(helperMsg);
        return false;
      }
      break;
    case 31:
      if(m == 2 || m == 4 || m == 6 || m == 9 || m == 11){
      //2�B4�B6�B9�B11��S��31��
        alert(helperMsg);
        return false;
      }
      break;
    default:

  }
 //alert(dateStr);
  return true;
}
//=============================================================

</script>

<script type="text/javascript">
//function exeadd(){
//   if ( document.forms["form1"]["exemen_w"].value == "clear" ){
//      document.forms["form1"]["wk_exe"].value = "";
//       }else{
//       if(document.forms["form1"]["wk_exe"].value == ""){
//         document.forms["form1"]["wk_exe"].value = document.forms["form1"]["exemen_w"].value;
//       }else{
//       document.forms["form1"]["wk_exe"].value = document.forms["form1"]["wk_exe"].value + "," + document.forms["form1"]["exemen_w"].value ;
//       }
//   }
//   	document.forms["form1"]["exemen_w"].value="" ;
//}
function exeadd(){
   if ( document.forms["form1"]["exemen_w"].value == "clear" ){
      document.forms["form1"]["wk_exe"].value = "";
       }else{
       if (document.forms["form1"]["wk_exe"].value == ""){
         document.forms["form1"]["wk_exe"].value = document.forms["form1"]["exemen_w"].value;
       }else{
         var p_str=document.forms["form1"]["wk_exe"].value;
         var p_n=p_str.search(document.forms["form1"]["exemen_w"].value);
         if (p_n>=0){
            var p_str1 = p_str.replace(document.forms["form1"]["exemen_w"].value,"");
            var pk_str = p_str1.replace(",,",",");
               document.forms["form1"]["wk_exe"].value = clearcoms(pk_str);
         }else{
            document.forms["form1"]["wk_exe"].value = document.forms["form1"]["wk_exe"].value + "," + document.forms["form1"]["exemen_w"].value ;
         }
       }
   }
   	document.forms["form1"]["exemen_w"].value="" ;
}
function clearcoms(pp_str){
   if (pp_str.charAt(0)==","){
      pp_str=pp_str.substr(1,pp_str.length-1);
   }
   if (pp_str.charAt(pp_str.length-1)==","){
      pp_str=pp_str.substr(0,pp_str.length-1);
   }
//   alert("pp_str="+pp_str) ;
   return(pp_str) ;
}

function showeagle(){
   wkrn="" ;   
   var str_gp1="<%=stra_gp1%>";
   txt="" ;
   chkwkr=document.forms["form1"]["all_worker"];
   for (i=0;i< chkwkr.length;++ i)
   {
      wkrn = chkwkr[i].value;
      if (str_gp1.search(wkrn) >= 0){
         chkwkr[i].checked=true;
       txt=txt + wkrn + " ";
     } else {
         chkwkr[i].checked=false;
      }
   }
//   document.getElementById("t_content").value="��ܡG " + txt ;
   return (true);
}
</script>
</center>
</body>
</html>
