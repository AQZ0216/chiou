<%@ Language=VBScript CODEPAGE=950 %>
<%
'Ū��temp-daywork.mdb���  '==================================
'pa_undo_date1=Request("undo_date1")         '���i���
'pa_doing_date1=Request("doing_date1")       '������
'pa_wk_class=Request("wk_class")                   '�u�@����
'pa_wk_group=Request("wk_group")                '�u�@�s��
'pa_wk_item=Request("wk_item")                     '�D��
'pa_wk_content=Request("wk_content")         '���e
'pa_wk_order=Request("wk_order")                 '���i��
'pa_all_worker=Request("all_worker")     '���|�H��
dim pa_undo_date1()   '���i���
dim pa_doing_date1()      '������
dim pa_wk_class()                  '�u�@����
dim pa_wk_group()               '�u�@�s��
dim pa_wk_item()                    '�D��
dim pa_wk_content()         '���e
dim pa_wk_order()                '���i��
dim pa_wk_doer()                '�u�@�H�����|�H��
dim pa_wk_undoer()                '�������u�@�H��
dim pa_wk_exe()                '����H��

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
	strSQL_show="Select * from " & tb_name & " where ipt_ok = 0"
	rstObj1.open strSQL_show,conDB,3,3
	'�p�����`��	
	totalput=rstObj1.recordcount
	if totalput=0 then
	else
redim pa_undo_date1(totalput)   '���i���
redim pa_doing_date1(totalput)      '������
redim pa_wk_class(totalput)                  '�u�@����
redim pa_wk_group(totalput)               '�u�@�s��
redim pa_wk_item(totalput)                    '�D��
redim pa_wk_content(totalput)         '���e
redim pa_wk_order(totalput)                '���i��
redim pa_wk_doer(totalput)                '�u�@�H�����|�H��
redim pa_wk_undoer(totalput)                '�������u�@�H��
redim pa_wk_exe(totalput)                '����H��
		'�C�X��ƶ���
		rstobj1.MoveFirst
		for j=1 to totalput
			pa_undo_date1(j-1)=rstObj1.fields("undo_date1")         '���i���
			pa_doing_date1(j-1)=rstObj1.fields("doing_date1")         '������
			pa_wk_class(j-1)=rstObj1.fields("wk_class")         '�u�@����
			pa_wk_group(j-1)=rstObj1.fields("wk_group")         '�u�@�s��
			pa_wk_item(j-1)=rstObj1.fields("wk_item")         '�D��
			pa_wk_content(j-1)=rstObj1.fields("wk_content")         '���e
			pa_wk_order(j-1)=rstObj1.fields("wk_order")         '���i��
			pa_wk_doer(j-1)=rstObj1.fields("wk_doer")         '�u�@�H�����|�H��
			pa_wk_undoer(j-1)=rstObj1.fields("wk_undoer")         '�������u�@�H��
			pa_wk_exe(j-1)=rstObj1.fields("wk_exe")         '����H��
			'����U�@���O��
			rstObj1.MoveNext
			if rstObj1.EOF=True then exit for
		next
	end if
	'������ƶ�
	rstObj1.Close
	'���]����ܼ� 
	set rstObj1=Nothing
	'������Ʈw
	conDB.Close
	'���]�����ܼ�
	set conDB=Nothing
'==================================

if totalput>0 then              '==================================
   ' �s��Access��Ʈwdaywork.mdb
   DBpath=Server.MapPath("./database/daywork.mdb")
   strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
   '�إ߸�Ʈw�s������
   set conDB= Server.CreateObject("ADODB.Connection")
   '�s����Ʈw	
   conDB.Open strCon
   '�}�Ҹ�ƪ�W��
   tb_name="work_data"

      for kj=1 to totalput          '==================================
         p_undo_date1=pa_undo_date1(kj-1)
         p_doing_date1=pa_doing_date1(kj-1)
         p_wk_class=pa_wk_class(kj-1)
         p_wk_group=pa_wk_group(kj-1)
         p_wk_item=pa_wk_item(kj-1)
         'p_wk_content=pa_wk_content(kj-1)
         p_wk_content=pa_wk_content(kj-1) & chr(13) & date()& "����s�W�C" 
         p_wk_order=pa_wk_order(kj-1)
         p_wk_doer=pa_wk_doer(kj-1)
         p_wk_undoer=pa_wk_undoer(kj-1)
         p_wk_exe=pa_wk_exe(kj-1)
         '�s�W��Ƥ�SQL���O�r��
         strSQL_add="Insert into "&tb_name&" ("
         strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
         strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,wk_exe,"
         strSQL_add=strSQL_add & "wk_doer,wk_undoer) values ('"
         strSQL_add=strSQL_add & p_undo_date1 &"','"
         strSQL_add=strSQL_add & p_doing_date1 &"','"
         strSQL_add=strSQL_add & p_wk_class &"','"
         strSQL_add=strSQL_add & p_wk_group &"','"
         strSQL_add=strSQL_add & p_wk_item &"','"
         strSQL_add=strSQL_add & p_wk_content &"','"
         strSQL_add=strSQL_add & p_wk_order &"','"
         strSQL_add=strSQL_add & p_wk_exe &"','"
         strSQL_add=strSQL_add & p_wk_doer&"','"
         strSQL_add=strSQL_add & p_wk_undoer&"')"
         '����SQL���O
         conDB.Execute strSQL_add
      next                        '==================================
   '������Ʈw
   conDB.Close
   '���]�����ܼ�
   set conDB=Nothing
end if                            '==================================

%>
<%
'=========�R���w�P�B�����==========================
'Ū��temp-daywork.mdb���  '==================================
dim pa_delwkid()   '�R�����u�@id

	' �s��Access��Ʈwtemp-daywork.mdb
	DBpath=Server.MapPath("./database/temp-daywork.mdb")
	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'�إ߸�Ʈw�s������
	set conDB= Server.CreateObject("ADODB.Connection")
	'�s����Ʈw	
	conDB.Open strCon
	'�}�Ҹ�ƪ�W��
	tb_name="del_work_data"
	'�إ߸�Ʈw�s������	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " where ipt_ok = 0"
	rstObj1.open strSQL_show,conDB,3,3
	'�p�����`��	
	totalput=rstObj1.recordcount
	if totalput=0 then
	else
	redim pa_delwkid(totalput)   '�R�����u�@id
		'�C�X��ƶ���
		rstobj1.MoveFirst
		for j=1 to totalput
			pa_delwkid(j-1)=rstObj1.fields("tmp_id")         '�R�����u�@id
			'����U�@���O��
			rstObj1.MoveNext
			if rstObj1.EOF=True then exit for
		next
	end if
	'������ƶ�
	rstObj1.Close
	'���]����ܼ� 
	set rstObj1=Nothing
	'������Ʈw
	conDB.Close
	'���]�����ܼ�
	set conDB=Nothing
'==================================
if totalput>0 then              '==================================
   ' �s��Access��Ʈwdaywork.mdb
   DBpath=Server.MapPath("./database/daywork.mdb")
   strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
   '�إ߸�Ʈw�s������
   set conDB= Server.CreateObject("ADODB.Connection")
   '�s����Ʈw	
   conDB.Open strCon
   '�}�Ҹ�ƪ�W��
   tb_name="work_data"
      for kj=1 to totalput          '==================================
			 	'�إ߸�Ʈw�s������	
				set rstObj1=Server.CreateObject("ADODB.Recordset")
				strSQL_show="Select * from " & tb_name & " where wk_id =" & pa_delwkid(j-1)
				rstObj1.open strSQL_show,conDB,3,3  	
					'�p�����`��	
					jt=rstObj1.recordcount
				'������ƶ�
				rstObj1.Close
				'���]����ܼ� 
				set rstObj1=Nothing
				if jt=1 then				      	
					'�R����Ƥ�SQL���O�r��
					strSQL_del="Delete from " & tb_name & " where wk_id =" & pa_delwkid(j-1)
					'����SQL���O
					conDB.Execute strSQL_del
				end if
      next                        '==================================
   '������Ʈw
   conDB.Close
   '���]�����ܼ�
   set conDB=Nothing
end if                            '==================================


%>
<html>
<head>
<title>�N�Ȧs�u�@�s�J�u�@��Ʈw��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--Ū�J�ù���ܼ˪O�� base_screen_�@��.css �ΦC�L�˪O�� base_print_�@��.css  -->
	<link rel="stylesheet" type="text/css" 
		media="screen" href="./css/base_screen.css" title="style_screen">
<!--�]�w�˪O�榡-->
<style type="text/css">
	<!--

	-->
</style>
</head>
<body>
<%
'�۰���������
Response.Write "<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> "
%>
</body>
</html>