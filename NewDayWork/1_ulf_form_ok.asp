<%@ Language=VBScript CODEPAGE=950 %>
<%
	'Ū���H���m�W
	worker = Session("worker")
%>
<%
'BASP21.DLL�N�ɮפW���{��
'�Х�����w��RegSvr32 Basp21.dll
'�i�N��椤��text�]��X��i�}�C�A����response.write �ܼơA�N�i�Hprint�X�ӤF
'-------------------------------------------------------------------
%>

<HTML> 
<HEAD>
<Title>�W���ɮץ\��{��</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
</HEAD>
<BODY>
<center>
<%
dim Upload,A,B,Text,Image,ImgName,ImgPath,RC 
'// Upload��BASP21����ϥΤ��ܼ� 
'// A���N���o���Ⱦ֦��h��Bytes���ܼ� 
'// B���N���o�����ର�G�i��X���ܼ� 
'// Text��Form�������ȱo�ܼ�
'// Image�����o�ӷ����ɵ�����|(�t�ɦW)���ܼ� 
'// ImgName���ӷ������ɦW���ܼ� 
'// ImgPath�����ɱ��x�s�����|���ɦW���ܼ�
'// RC���ˬd�ɮ׬O�_�W�Ǥ��ܼ� 
set Upload = server.CreateObject("BASP21") '// �إ�BASP21���A������ 
A = request.TotalBytes '// �N���o���ȱoBytes�� 
B = request.BinaryRead(A) '// �N���o�����ର�G�i��X
Text= Upload.Form(B,"text") '// ���oForm�Ȩ��ର�G�i��X    '�u�@wk_id
p_item= Upload.Form(B,"item") '// ���oForm�Ȩ��ର�G�i��X  '�ɮ׻���
Image = Upload.FormFileName(B,"image") '// ���o���ɵ�����|(�t�ɦW)     '�ɮצW��
ImgName = mid(Image,InStrRev(Image,"\")+1) '// ���o���ɦW��(�t���ɦW) \
file= text &"_"& ImgName     '�ɮצW�٧אּ wk_id+���ɦW

response.Write "ImgName="& ImgName &"<br>"   '�ڼg
response.Write "wk_id="& text &"<br>"  '�ڼg

'ImgPath = server.MapPath("addfile") & "\" & ImgName '//�A���x�s�ɮר���̪����|���ɦW
ImgPath = server.MapPath("file_att") & "\" & file '//�A���x�s�ɮר���̪����|���ɦW

'// �ˬd���A�����w�����|�O�_���ۦP���ɮ�
if Upload.FileCheck(ImgPath) >= 0 then 
     'set Upload = nothing '//�M�Ū���[���G�ȥ��M�šA�_�h�q�X�Ӫ������|�ܦ��G�i��X�C]
      'response.Write("") &vbcrlf 
      'response.End
      Response.Write "���A�������ۦP�ɦW <br>"
      old_file=1
else '// �Y�ɮפ��s�b
      'RC = Upload.FormSaveAs(B,"image",ImgPath) '//�W���ɮױqForm���ɮת��image����A����ImgPath
      '// �ˬd�ɮפW�Ǧ��\�P�_
      Response.Write "���A�����S���ۦP�ɦW <br>"
      old_file=0
end if 

'response.Write "ImgPath="& ImgPath &"<br>"   '�ڼg
'response.end
'============������ɦW=====================20111118
str_except_file="avi�Bmpg�Bmlv�Bmpe�Bmpeg�Basf�Bwmv�B.rm�Brmvb"   '�ҥ~���ɦW
'file_ext=right(ImgName,InStrRev(ImgName,".",-1,1)-1)
file_ext=right(ImgName,3)
'response.write "���ɦW�G"& file_ext &"<br>"
if instr(1,str_except_file,file_ext,1)=0 then
   RC = Upload.FormSaveAs(B,"image",ImgPath)      '//�W���ɮױqForm���ɮת��image����A����ImgPath
else
   RC=0
end if
'============������ɦW=====================20111118
'RC = Upload.FormSaveAs(B,"image",ImgPath)      '//�W���ɮױqForm���ɮת��image����A����ImgPath

set Upload = nothing

if RC > 0 then
      Response.Write "[ "&RC&" ] byte�W�Ǧ��\ .<br>"
      Response.Write "�ɮפW�Ǧ��\ .<br>"
      '============= �N�W���ɮ׸�ƿ�J��Ʈw�� ==================
      p_wk_id=text     'wk_id
      p_fl_name=file   '�ɮצW��
      p_fl_size=RC      '�ɮפj�p
      p_fl_date=date() '���ɤ��
      p_fl_item=p_item '�ɮ׻���
      p_fl_inputer=worker '�W���ɮפH��

    if old_file=0 then
         p_fl_history=now()&"�e"&worker&"�f�W���ɮסC"
         ' �s��Access��Ʈwattach_file.mdb
         DBpath_fl=Server.MapPath("./database/attach_file.mdb")
         strCon_fl="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fl
         '�إ߸�Ʈw�s������
         set conDB_fl= Server.CreateObject("ADODB.Connection")
         '�s����Ʈw	
         conDB_fl.Open strCon_fl
         '�}�Ҹ�ƪ�W��
         tb_name_fl="file_data"
         '�s�W��Ƥ�SQL���O�r��
         strSQL_add_fl="Insert into "&tb_name_fl&" ("
         strSQL_add_fl=strSQL_add_fl & "wk_id,fl_name,fl_size,fl_date,fl_item,fl_inputer,fl_history) values ('"
         strSQL_add_fl=strSQL_add_fl & p_wk_id &"','"
         strSQL_add_fl=strSQL_add_fl & p_fl_name &"','"
         strSQL_add_fl=strSQL_add_fl & p_fl_size &"',#"
         strSQL_add_fl=strSQL_add_fl & p_fl_date &"#,'"
         strSQL_add_fl=strSQL_add_fl & p_fl_item &"','"
         strSQL_add_fl=strSQL_add_fl & p_fl_inputer &"','"
         strSQL_add_fl=strSQL_add_fl & p_fl_history &"')"
         '����SQL���O
         conDB_fl.Execute strSQL_add_fl
         '������Ʈw
         conDB_fl.Close
         '���]�����ܼ�
         set conDB_fl=Nothing
   else
         p_fl_history=now()&"�e"&worker&"�f�W���ɮר��N���ɮסC"
         ' �s��Access��Ʈwattach_file.mdb
         DBpath_fl=Server.MapPath("./database/attach_file.mdb")
         strCon_fl="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fl
         '�إ߸�Ʈw�s������
         set conDB_fl= Server.CreateObject("ADODB.Connection")
         '�s����Ʈw	
         conDB_fl.Open strCon_fl
         '�}�Ҹ�ƪ�W��
         tb_name_fl="file_data"
         '�إ߸�Ʈw�s������	
         set rstObj1_fl=Server.CreateObject("ADODB.Recordset")
         strSQL_show_fl="Select * from " & tb_name_fl & " where wk_id="& p_wk_id & " and fl_name like '"& p_fl_name &"' order by fl_name asc"
         rstObj1_fl.open strSQL_show_fl,conDB_fl,1,3
         rstObj1_fl.fields("fl_size")=p_fl_size
         rstObj1_fl.fields("fl_date")=p_fl_date
         rstObj1_fl.fields("fl_item")=p_fl_item
         rstObj1_fl.fields("fl_inputer")=p_fl_inputer
         rstObj1_fl.fields("fl_history")=rstObj1_fl.fields("fl_history") & chr(13) & p_fl_history
         rstObj1_fl.UpdateBatch
         '������ƶ�
         rstObj1_fl.Close
         '���]����ܼ� 
         set rstObj1_fl=Nothing
         '������Ʈw
         conDB_fl.Close
         '���]�����ܼ�
         set conDB_fl=Nothing
   end if
      '============= �N�W���ɮ׸�ƿ�J��Ʈw�� ==================
      set Upload = nothing

   myURL="wk_show.asp?wk_id="& text
   Response.Redirect (myURL)
else
      set Upload = nothing '//�M�Ū���[���G�ȥ��M�šA�_�h�q�X�Ӫ������|�ܦ��G�i��X�C]
      Response.Write "[ "&RC&" ] byte�W�� .<br>"
      'response.Write("") &vbcrlf
      'response.End
      'Response.Write "�ɮפW�ǥ��� !<br>"
%>
   <%=ImgName%> �ɮפW�ǥ��� !<br>
   ���ɦW��"avi�Bmpg�Bmlv�Bmpe�Bmpeg�Basf�Bwmv�Brm�Brmvb"�A�L�k�W�ǡC
<hr>
<a href="wk_show.asp?wk_id=<%=text%>" target="_self">�^�u�@����</a>

<%
end if
%>

</center>
</BODY>
</HTML>