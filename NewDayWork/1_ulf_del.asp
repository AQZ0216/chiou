<%@ Language=VBScript CODEPAGE=950 %>
<%
'����ˬd�u�@�O�_�s�b
function exist_wkid(pwk_id)
      ' �s��Access��Ʈwdaywork.mdb
      DBpath_fe=Server.MapPath("./database/daywork.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '�إ߸�Ʈw�s������
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB_fe.Open strCon_fe
      '�}�Ҹ�ƪ�W��
      tb_name_fe="work_data"
      '�إ߸�Ʈw�s������	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '������ƶ�
      rstObj1_fe.Close
      '���]����ܼ�
      set rstObj1_fe=Nothing
      '������Ʈw 
      conDB_fe.Close
      '���]�����ܼ�
      set conDB_fe=Nothing
      exist_wkid=totalput_fe
end function
%>
<%
' ------------------------------------------
' �N�ɮײ���file_del�ؿ�
Sub Movefile(strFile)
   'strFile �ɮצW��
   strDir1=Server.MapPath("./file_att")    '�� �ε������|���o�ɮצ�m
   strDir2=Server.MapPath("./file_del")   '�s �ε������|���o�ɮצ�m
   response.write strFile &"<br>"
'   response.end
      '�ŧi����objFSO objInStream���ܼ�intCount strFileContent strInLine
	Dim objFSO, objInStream, intCount, strFileContent, strInLine
	'�]�w�ɮצs������
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    '�R���¥ؿ������ɮ�	
	if  (objFSO.FileExists(strDir2 & "\" & strFile)) then
		objFSO.DeleteFile(strDir2 & "\" & strFile)
	else
	end if
	'�N�ɮײ��ܷs�ؿ�
	Set MyFile = objFSO.GetFile(strDir1 & "\" & strFile)
	MyFile.Move Server.MapPath("./file_del")& "\"
    Set objFSO = Nothing
    response.write "<hr>"
end sub 
' ------------------------------------------
%>
<%
	'Ū���H���m�W
	worker = Session("worker")
	fl_id=Request("fl_id")
%>
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
strSQL_show="Select * from " & tb_name & " where fl_id =" & fl_id &" and del_ok = false"
rstObj1.open strSQL_show,conDB,3,3
totalput=rstObj1.recordcount
if totalput=0 then
else
	'�C�X��ƶ���
	rstobj1.MoveFirst
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")      '�ɮצW��
		pfl_item=rstObj1.fields("fl_item")           '�ɮ׻���
		pfl_date=rstObj1.fields("fl_date")           '���ɤ��
         rstObj1.fields("del_ok") = true                  '�O�_�R��
         Movefile pfl_name                   '�����ɮצ�m
         rstObj1.UpdateBatch
end if
'������ƶ�
rstObj1.Close
'���]����ܼ�
set rstObj1=Nothing

'==============�R�����============================
'�R����Ƥ�SQL���O�r��
'strSQL_del="Delete from " & tb_name & " where fl_id =" & fl_id
'����SQL���O
'conDB.Execute strSQL_del
'==============�R�����============================
'������Ʈw 
conDB.Close
'���]�����ܼ�
set conDB=Nothing

if exist_wkid(pwk_id)=1 then
   'response.write "�ɮקR������"
   myURL="wk_show.asp?wk_id="&pwk_id
   Response.Redirect (myURL)
else
   'response.write "�ɮקR������"
   myURL="1_ulf_list.asp"
   Response.Redirect (myURL)
end if
%>

<HTML> 
<HEAD>
<Title>�W���ɮץ\��{��</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
</HEAD>
<BODY>
<center>
�ɮקR������!!
<hr>
<a href="wk_show.asp?wk_id=<%=pwk_id%>" target="_self">�^�u�@����</a>
</center>
</BODY>
</HTML>