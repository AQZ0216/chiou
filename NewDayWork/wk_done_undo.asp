<%@ Language=VBScript CODEPAGE=950 %>
<%
'20120517��s ========
   'Ū���H���m�W
   worker = Session("worker")
   wk_id=Request("wk_id")

%>
<html>
<head>
<title>��ƭק�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'##FFFFcc'}
--></style>
</head>
<body>
<center>
<!-- �}�Ҹ�Ʈw -->
<!-- #Include file = "./include/opendb_daywork.inc" -->
<%
'�إ߸�Ʈw�s������  
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,3 
'Ū�����

wk_content=rstObj1.fields("wk_content")
wk_doer=rstObj1.fields("wk_doer")                    '�u�@�H��
wk_checker=rstObj1.fields("wk_checker")           '�w�ˬd�H��(�Ҧ��u�@�H��)
wk_undoer=rstObj1.fields("wk_undoer")               '�������u�@�H��
wk_finisher=trim(rstObj1.fields("wk_finisher"))      '�����H��(��g�������H��)

'�H�W�b�����W�椤����
if isnull(wk_finisher) or wk_finisher="" then
   '�H�W�b�����W�椤����
   wk_checker=replace(wk_checker,worker,"")
   wk_checker=replace(wk_checker,",,",",")
   if left(wk_checker,1)="," then
      wk_checker=replace(wk_checker,",","",1,1)
   end if
else
   '�H�W�b�����W�椤����
   wk_checker=replace(wk_checker,worker,"")
   wk_checker=replace(wk_checker,",,",",")
   if left(wk_checker,1)="," then
      wk_checker=replace(wk_checker,",","",1,1)
   end if
end if

'�N�H�W�b�������u�@�̤��W��[�J
if wk_undoer="" or isnull(wk_undoer) then
   wk_undoer=worker
else
   wk_undoer=worker & "," & wk_undoer 
end if
'�b�u�@���e���W�[��������ΤH�W
wk_content=wk_content & chr(13) & worker & "��" & date() &"���������u�@"

rstObj1.fields("wk_content")=wk_content
rstObj1.fields("done_date1")=done_date1
rstObj1.fields("wk_checker")=wk_checker
rstObj1.fields("wk_finisher")=wk_finisher
rstObj1.fields("wk_undoer")=wk_undoer
rstObj1.UpdateBatch


'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
%>

<%
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing

'strbackURL=session("strbackURL")
strbackURL="wk_show.asp?wk_id="&wk_id
response.redirect(strbackURL)

%>
</body>
</html>
