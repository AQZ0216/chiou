<% @codepage=950%>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<!-- #Include file = "./include/array_wkclass.inc" -->
<!-- #Include file = "./include/array_wkgroup.inc" -->
<!-- #Include file = "./include/workinput.inc" -->
<!-- #Include file = "./misc_data/array_place.inc" -->	
<!-- #Include file = "./misc_data/array_thing.inc" -->
<!-- #Include file = "./include/array_pjn.inc" -->
<%
'Ū���K�X���
Function find_crewpwd(p_wkr)
      ' �s��Access��Ʈw../holiday/database/crew.mdb
      DBpath_p=Server.MapPath("../holiday/database/crew.mdb")
      strCon_p="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_p
      '�إ߸�Ʈw�s������
      set conDB_p= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB_p.Open strCon_p
      '�}�Ҹ�ƪ�W��
      tb_name_p="crew"
      '�إ߸�Ʈw�s������	
      set rstObj1_p=Server.CreateObject("ADODB.Recordset")
      strSQL_show_p="Select * from " & tb_name_p & " where worker like '" & p_wkr &"'"
      rstObj1_p.open strSQL_show_p,conDB_p,3,1
      totalp=rstObj1_p.recordcount
      if totalp=0 then
         p_pwd="0"
      else
         p_pwd=rstObj1_p.fields("wkr_pwd")		'�K�X
      end if
      '������ƶ�
      rstObj1_p.Close
      '���]����ܼ� 
      set rstObj1_p=Nothing
      '������Ʈw
      conDB_p.Close
      '���]�����ܼ� 
      set conDB_p=Nothing
      find_crewpwd=p_pwd
End Function
%>
<%
'---------------------------------------------------------------------------
wkr_pwd=session("wkr_pwd") 'Ū���K�X
if session("wkr_pwd")="" or isnull(wkr_pwd) then
      wkr_pwd=request("wkr_pwd") 'Ū���K�X
end if
'chk_str="���`�B���B����B���z�B�f�S"
chk_str=""
'if instr(1,chk_str,worker,1)>0 and instr(1,chk_str,worker_old,1)=0 then
if instr(1,chk_str,worker,1)>0 then
   if instr(1,chk_str,worker_old,1)>0 then
   'response.write "worker_old="&worker_old
   'response.end
   else
         'Ū����Ʈw�K�X
      '   response.write "worker="&worker&"<br>"
         db_pwd=find_crewpwd(worker)
         if db_pwd=wkr_pwd then
           session("wkr_pwd")=db_pwd
         else
            str_url="./0_login_pwd.asp?worker="&worker
            response.redirect(str_url)      '��}��K�X��J�e��
         end if
      end if
end if
'---------------------------------------------------------------------------
%>
<%
'Generates date in yyyy-mm-dd format
Function GetFormattedDate(setDate)
   strDate = CDate(setDate)
   strDay = DatePart("d", strDate)
   strMonth = DatePart("m", strDate)
   strYear = DatePart("yyyy", strDate)
   If strDay < 10 Then
     strDay = "0" & strDay
   End If
   If strMonth < 10 Then
     strMonth = "0" & strMonth
   End If
   GetFormattedDate = strYear & "-" & strMonth & "-" & strDay
End Function
%>
<%
	'Ū�����
	worker = request("worker") 'Ū���H��
	datecode=request("datecode")'Ū�����
%>
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<title>�i<%=worker%>�j�u�@--�s�W</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="./css/w3-cht.css">
<link rel="stylesheet" href="./css/font-awesome.min.css">

<style>

</style>
</head>
<body class="vt-container">

<!-- ���Y�}�l -->
<!-- #Include file = "./include/zec-header_r1.inc" -->
<!-- ���Y���� -->

<!-- ����}�l -->
<div class="vt-container w3-pale-red w3-center" >
   <div class="w3-row w3-center " ><!-- ���e start -->
<!-- ----------------------------------------------------------------- ���estart -----------------------------------------------------------------  -->   
<form name="form1" action="zec-work_add_ok_r1.asp" method="post" >
      <table class="w3-table-all" ><!-- �\��� start -->
         <col style="width:100%;background-color:#ffdddd;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <td style="text-align:center;height:55px;">
            <button class="w3-button w3-white w3-xlarge " style="padding:2px;margin:0px;" >�i<%=worker%>�j��@�u�@���طs�W</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�^�W�@��" onclick="javascript:history.go(-2)">�^�W�@��</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�T�w�s�W" onclick="">�T�w�s�W</button>
            <input type="reset" class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="���]���" value="���]���" >
           </td>
         </tr>
      </table><!-- �\��� end -->
      
<div class="w3-container w3-center" style="margin:0px;padding:0px;height:520px;overflow:auto;"><!-- �u�@���ت� start -->
      <div class="w3-responsive" ><!-- div w3-responsive start -->
      <table class="w3-table-all" >
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <th style="text-align:center;background-color:#ffdddd;border:1px solid #000;">�u�@�s��</th>
           <th style="text-align:center;background-color:#ddffdd;border:1px solid #000;">�M�צW��</th>
           <th style="text-align:center;background-color:#ffdddd;border:1px solid #000;">�u�@�s��</th>
           <th style="text-align:center;background-color:#ddffdd;border:1px solid #000;">�u�@����</th>
           <th style="text-align:center;background-color:#ffdddd;border:1px solid #000;">���i��</th>
           <th style="text-align:center;background-color:#ddffdd;border:1px solid #000;">���i���</th>
           <th style="text-align:center;background-color:#ffdddd;border:1px solid #000;">������</th>
         </tr> 
         <tr style="border:1px solid #000;" >
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;padding:0px;margin:0px;"><!-- �u�@�s�� -->
            	<select name="wk_group" style="width:100%;padding:0px;margin:0px;height:100%;">
            <%
            		response.write "<option value='"&wk_group_a(0)&"' selected>"&wk_group_a(0)
            	for i=2 to wk_group_no
            		response.write "<option value='"&wk_group_a(i-1)&"'>"&wk_group_a(i-1)
            	next
            %>
            	</select>
            </td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;padding:0px;margin:0px;"><!-- �M�צW�� -->
            	<select name="wk_pjn" style="width:100%;padding:0px;margin:0px;height:100%;" >
            <%
            		response.write "<option value='0' selected>"
            		'response.write "<option value='"&pjnid_a(0)&"' >"&pjn_a(0)
            	for i=1 to pjn_no
            		response.write "<option value='"&pjnid_a(i-1)&"�A"&pjn_a(i-1)&"'>"&pjn_a(i-1)
            	next
            %>
            	</select>
            </td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;padding:0px;margin:0px;"><!-- �u�@�s�� -->
               �۰ʽs��
            </td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;padding:0px;margin:0px;"><!-- �u�@���� -->
            	<select name="wk_class" style="width:100%;padding:0px;margin:0px;height:100%;" >
            <%
            	for i=1 to wk_class_no
            		response.write "<option value='"&wk_class_a(i-1)&"'>"&wk_class_a(i-1)
            		 if wk_class_a(i-1)="Z" then response.write "-���n����"
            	next
            %>
            	</select>
            </td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;padding:0px;margin:0px;"><!-- ���i�� -->
               <input type='text' name='wk_order' value='<%=worker%>' style="width:100%;" readonly>
            </td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;padding:0px;margin:0px;"><!-- ���i��� -->
               <input type="date" id="datePicker" name='undo_date1' value="<%=GetFormattedDate(datecode)%>" style="width:100%;" >
            </td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;padding:0px;margin:0px;"><!-- ������ -->
               <input type="date" name="doing_date1" value="<%=GetFormattedDate(date())%>" style="width:100%;">
            </td>
         </tr>
         <tr style="border:1px solid #000;" >
            <td>����H��</td>
            <td colspan="6" style="text-align:left;background-color:#ffdddd;border:1px solid #000;">
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
			<option value="<%=stra_dp10%>" >�ڮa�A�~</option>
			<option value="<%=stra_dpa1%>" >�~�@</option>
			<option value="<%=stra_dpa2%>" >�~�G</option>
			<option value="<%=stra_dpa3%>" >�~Three</option>
			<option value="<%=stra_dpa4%>" >YES</option>
			<option value="<%=stra_dpa5%>" >�_�w�~�K</option>
		</SELECT>	
				(�п�J����ѻP�H��)	
            </td>
         </tr>
         <tr style="border:1px solid #000;" >
            <td>�X�u�H��</td>
            <td colspan="6"><%=wk_att%></td>
         </tr>
         <tr style="border:1px solid #000;" >
            <td>���|�H��</td>
            <td colspan="6"><%=wk_doer%></td>
         </tr>
         <tr style="border:1px solid #000;" >
         	<td>�����H��</td>
         	<td colspan="6"><%=wk_checker%></td>
         </tr>
         <tr style="border:1px solid #000;" >
         	<td>�������H��</td>
         	<td colspan="6"><%=wk_undoer%></td>
         </tr>
         <tr style="border:1px solid #000;" >
         	<td align="right">�D���G</td>
         	<td colspan="6"><%=wk_item%></td>
         </tr>
         <tr style="border:1px solid #000;" >
         	<td align="right" valign="top">���椺�e�G</td>
         	<td colspan="6" style="padding:0px;margin:0px;">
               <% 
                  pp_wk_content=replace(wk_content,chr(13),"<br>",1,-1,1)
                  'response.write pp_wk_content
               %>
               <div style="margin:0px;padding:0px;height:200px;overflow:auto;">
               <%=pp_wk_content%>
               </div>
          	</td>
         </tr>
         <tr style="border:1px solid #000;" data-ng-bind="">
         	<td align="right"><font color="red">�[�K��r�G</font></td>
         	<td colspan="6"><%=pwk_password%></td>
         </tr>
      </table>     
      </div><!-- div w3-responsive end -->
     
</div><!-- �u�@���ت� end -->
</form>
<!-- -----------------------------------------------------------------���e end----------------------------------------------------------------- -->   
   </div><!-- ���e end -->
</div>
<!-- ���嵲�� -->
<!-- �����}�l -->
<!-- #Include file = "./include/zec-footer_r1.inc" -->
<!-- �������� -->

<script language="JavaScript">
    function url_new(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        //window.location.href = pp_url; //�쭶����s
        window.open(pp_url) ; //�}�ҷs����
        return true;
    }   
    function url_show(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        window.location.href = pp_url; //�쭶����s
        //window.open(pp_url) ; //�}�ҷs����
        return true;
    }
    function url_show_confirm(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        //
        aws=confirm("�T�w����i"+pp_url+"�j�ܡH�H");
        if (aws)
        {
            window.location.href = pp_url; //�쭶����s
            //window.open(pp_url) ; //�}�ҷs����
            return true;         
        }
        //else
        //{
        // alert("��������i"+pp_url+"�j�I�I");
        //}       
    }       
    function content_show(pp_url){
        var iframe1=document.getElementById("ifrm_content");
        iframe1.src=pp_url;
        return true;
    }    

</script>

</body>
</html>