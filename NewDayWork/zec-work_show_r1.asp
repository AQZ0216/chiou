<% @codepage=950%>
<!-- #Include file = "./include/array_worker_crew.inc" -->
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
	'Ū���H���m�W
	worker = Request("worker")
	wk_id=Request("wk_id")
   wk_chk=Request("wk_chk")
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
undo_date1=rstObj1.fields("undo_date1")
doing_date1=rstObj1.fields("doing_date1")
done_date1=rstObj1.fields("done_date1")
p_wkid=rstObj1.fields("wk_id")
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
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<title>�i<%=worker%>�j�u�@�޲z</title>
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
   
      <table class="w3-table-all" ><!-- �\��� start -->
         <col style="width:100%;background-color:#ffdddd;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <td style="text-align:center;height:55px;">
<% 
if isnull(wk_headline) then wk_headline=0
if cint(wk_headline) > 5 then 
%>
<img src="./img/gnome_chess.png" width=32 onclick="parent.content.location.href='0_wk_headline_off_20140728.asp?wk_id=<%=wk_id%>'" title="�w�b�]���O�T�����A���X�]���O�T��">
<% else %>
<img src="./img/gnome_chess_d.png" width=32 onclick="parent.content.location.href='0_wk_headline_on_20140728.asp?wk_id=<%=wk_id%>'" title="���b�]���O�T�����A��J�]���O�T��">
<% end if %>           
            <button class="w3-button w3-white w3-xlarge " style="padding:2px;margin:0px;" >��@�u�@�������</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�^�W�@��" onclick="javascript:history.go(-1)">�^�W�@��</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�ק�u�@" onclick="url_show_confirm('zec-work_edit_r1.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">�ק�u�@</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�R���u�@" onclick="url_show_confirm('zec-work_del_r1.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">�R���u�@</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�����u�@" onclick="url_show_confirm('zec-work_del_r1.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">�����u�@</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="���s���i" onclick="url_show_confirm('zec-work_readd.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">���s���i</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�C�L���e" onclick="url_open('zec-work_print_si.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">�C�L���e</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�ର�M��" onclick="url_show_confirm('zec-work_gpchg_special.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">�ର�M��</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�ର�@��" onclick="url_show_confirm('zec-work_gpchg_normal.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">�ର�@��</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="�W�Ǫ���" onclick="url_show_confirm('zec-1_ulf_form.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">�W�Ǫ���</button>
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
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;"><%=wk_group%></td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;"><%=wk_pjn%></td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;"><%=wk_id%></td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;"><%=wk_class%></td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;"><%=wk_order%></td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;"><%=undo_date1%></td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;"><%=doing_date1%></td>
         </tr>
         <tr style="border:1px solid #000;" >
            <td>����H��</td>
            <td colspan="6"><%=wk_exe%></td>
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
      <div class="w3-responsive" ><!-- div w3-responsive start -->
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
<table class="w3-table-all" >
<col style="width:80px;border:1px solid #000;">
<col style="border:1px solid #000;">
<col style="width:300px;border:1px solid #000;">
<col style="width:90px;border:1px solid #000;">
<col style="width:150px;border:1px solid #000;">
<tr style="border:1px solid #000;" >
<td colspan=5>����C��</td>
</tr>
<tr style="background-color:#ffdddd;border:1px solid #000;">
<th>�Ǹ�</th>
<th>�ɮ׻���</th>
<th>�ɮצW��  [�W�Ǫ�]</th>
<th>���ɤ��</th>
<th>�\��</th>
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
<tr style="border:1px solid #000;">
<td style="text-align:center;"><%=fi%></td>
<td >
<a href="./zec-1_ulf_item_edit.asp?worker=<%=worker%>&wk_id=<%=pwk_id%>&fl_id=<%=pfl_id%>" target="_self" title="�ק��ɮ׻����C" ><img src="./img/change.png" style="vertical-align:middle;height:16px;cursor:hand;border:0;" ></a>
<%=pfl_item%>
</td>
<td ><a href="./file_att/<%=pfl_name%>" target="_blank" title="<%=pfl_history%>"><%=str_pfl_name%></a>  [<%=pfl_inputer%>]</td>
<td ><%=pfl_date%></td>
<td >
<button class="w3-button w3-blue" style="padding:2px;margin:0px;" onclick="url_show_confirm('zec-1_ulf_form.asp?worker=<%=worker%>&wk_id=<%=pwk_id%>')" title="�u�@���� [ wk_id=<%=pwk_id%> ] �s�W�ɮסC">�i�s�j</button>
<button class="w3-button w3-blue" style="padding:2px;margin:0px;" onclick="url_show_confirm('zec-1_ulf_del.asp?worker=<%=worker%>&wk_id=<%=pwk_id%>&fl_id=<%=pfl_id%>')" title="�N�ɮקR���C">�i�R�j</button>
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
      </div><!-- div w3-responsive end -->
      
</div><!-- �u�@���ت� end -->

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