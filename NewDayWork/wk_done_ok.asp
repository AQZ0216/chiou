<%@ Language=VBScript CODEPAGE=950 %>
<%
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
pdo_date=rstObj1.fields("doing_date1")
wk_content=rstObj1.fields("wk_content")
wk_doer=rstObj1.fields("wk_doer")
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_finisher=trim(rstObj1.fields("wk_finisher"))
'if isnull(wk_finisher) then
if isnull(wk_finisher) or wk_finisher="" then
   done_date1=cstr(date())
   wk_finisher=worker
   '�N�H�W�[�J�����u�@�̤��W�椤
   wk_checker=worker

else
   if isnull(rstObj1.fields("done_date1")) then
      done_date1=cstr(date())
   else
      done_date1=cstr(rstObj1.fields("done_date1"))
   end if
   '�N�H�W�[�J�����u�@�̤��W�椤
   wk_checker=wk_checker&","&worker
end if

'�N�H�W�b�������u�@�̤��W��h��
wk_undoer=replace(wk_undoer,worker,"")
wk_undoer=replace(wk_undoer,",,",",")
if left(wk_undoer,1)="," then
   wk_undoer=replace(wk_undoer,",","",1,1)
end if
'�b�u�@���e���W�[��������ΤH�W
wk_content=wk_content & chr(13) & worker & "��" & date() &"�����u�@"

'20100312��s ======== 
rstObj1.fields("wk_content")=wk_content
rstObj1.fields("done_date1")=done_date1
rstObj1.fields("wk_checker")=wk_checker
rstObj1.fields("wk_finisher")=wk_finisher
rstObj1.fields("wk_undoer")=wk_undoer
rstObj1.UpdateBatch
'20100312��s ======== 

'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
%>
<%
'20100312��s ======== 
'�ק��Ƥ�SQL���O�r�� �������
'strSQL_edit="Update "&tb_name&" set wk_content='"& wk_content &"'"
'strSQL_edit=strSQL_edit & ",done_date1=#"& done_date1 &"#"
'strSQL_edit=strSQL_edit & ",wk_checker='"& wk_checker &"'"
'strSQL_edit=strSQL_edit & ",wk_finisher='"& wk_finisher &"'"
'strSQL_edit=strSQL_edit & ",wk_undoer='"& wk_undoer &"'"
'strSQL_edit=strSQL_edit & " where wk_id =" & wk_id
'����SQL���O
'conDB.Execute strSQL_edit
'20100312��s ======== 
%>
<%
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing

'strbackURL=session("strbackURL")
nWeeksn =DatePart("ww",date()) 'Ū���g��
nWeeks =DatePart("ww",pdo_date) 'Ū���g��
nYear = Year(pdo_date)

if nWeeksn=nWeeks then
	strbackURL="wk_week_now.asp?nWeeks="&nWeeks&"&nYear="&nYear
else 
	strbackURL=session("strbackURL")
end if
response.redirect(strbackURL)

%>
<!-- <script language="Javascript">
   alert("��ƭק粒���I�I");
//   location.href="wk_lst_doing.asp";
   location.href="wk_Calendar_all.asp";
</script> -->

</body>
</html>
