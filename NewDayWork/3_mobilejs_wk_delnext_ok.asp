<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
'�P�_�O�_��J�u�@���� 
keyword=request("wk_class")
'if keyword="" then 
	'response.redirect("wk_add.asp")
'else
'end if
p_wk_id=Request("wk_id")         'id

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

'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & p_wk_id
rstObj1.open strSQL_show,conDB,3,1
'Ū�����
p_undo_date1=rstObj1.fields("undo_date1")
p_doing_date1=rstObj1.fields("doing_date1")
p_wk_class=rstObj1.fields("wk_class")
p_wk_group=rstObj1.fields("wk_group")
p_wk_item=rstObj1.fields("wk_item")
p_wk_content=rstObj1.fields("wk_content")
p_wk_order=rstObj1.fields("wk_order")
p_wk_exe=rstObj1.fields("wk_exe")
p_wk_doer=rstObj1.fields("wk_doer")
p_wk_undoer=rstObj1.fields("wk_undoer")

'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing

'�R����Ƥ�SQL���O�r��
strSQL_del="Delete from " & tb_name & " where wk_id =" & p_wk_id
'����SQL���O
conDB.Execute strSQL_del

'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing 
%>
<%
'===============�N�n�R������ưO����temp-daywork.mdb��===============
p_tmp_id=p_wk_id
p_iptok=0
' �s��Access��Ʈwtemp-daywork.mdb
DBpath=Server.MapPath("./database/temp-daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="del_work_data"

'�s�W��Ƥ�SQL���O�r��
strSQL_add="Insert into "&tb_name&" (tmp_id,"
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
strSQL_add=strSQL_add & "wk_doer,wk_undoer,ipt_ok) values ('"
strSQL_add=strSQL_add & p_tmp_id &"','"
strSQL_add=strSQL_add & p_undo_date1 &"','"
strSQL_add=strSQL_add & p_doing_date1 &"','"
strSQL_add=strSQL_add & p_wk_class &"','"
strSQL_add=strSQL_add & p_wk_group &"','"
strSQL_add=strSQL_add & p_wk_item&"','"
strSQL_add=strSQL_add & p_wk_content &"','"
strSQL_add=strSQL_add & p_wk_order &"','"
strSQL_add=strSQL_add & p_wk_exe &"','"
strSQL_add=strSQL_add & p_wk_doer&"','"
strSQL_add=strSQL_add & p_wk_undoer&"',"
strSQL_add=strSQL_add & p_iptok&")"

'����SQL���O
conDB.Execute strSQL_add

'������Ʈw
conDB.Close
'���]�����ܼ�
set conDB=Nothing
%>
<html>
<head>
<title>��Ƨ���s�W</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
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
		font-size:4.5mm; 			/*�r��j�p*/
		background-color:'#F0FFF0';
     }
td{
   margin:0 0 0 0;      /*��t�W�U���k*/
   border-color:'#808080'; /*���~���C��*/ 
   border-style:solid;     /*���~�ؽu��*/
   border-width:1px;    /*���~�ثp��*/  
   vertical-align:middle;  /*�r�髫������覡*/
   font-size:4.5mm;
   }
table{
   margin:0 0 0 0;      /*��t�W�U���k*/
   border-collapse:collapse;  /*��اΦ����X*/
   }
--></style>
</head>
<body>
<center>

<script language="Javascript">
	alert("��ƧR�������I�I");
	location.href="wk_Calendar_all.asp";
</script>
</body>
</html>
