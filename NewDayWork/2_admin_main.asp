<%@ Language=VBScript CODEPAGE=950 %>
<%
if not(session("admin_pwd")="28283939") then
      pwd=request("pwd")
      session("admin_pwd")=pwd
elseif request("logout")="1" then
      session("admin_pwd")=""
end if
%>
<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<%
if session("admin_pwd")="28283939" then
%>
<a href="1_ulf_list.asp" >�Ҧ��W���ɮפ��C��</a><br>
<a href="1_ulf_del_list.asp" >�w�R�����W���ɮצC��</a><br>
<a href="2_admin_main.asp?logout=1" title="�n�X��x�޲z">�n�X</a>
<hr>
<a href="wk_revise_order.asp" title="����קﬣ�u�̦W��">����קﬣ�u�̦W��</a>
<hr>
<a href="z_wk_lst_null.asp" title="���u�̬��ťդ��Ҧ��u�@�C��">���u�̬��ťդ��Ҧ��u�@�C��</a>
<hr>
<a href="z_wk_revise_order_null_ok.asp" title="����ק�ťլ��u�̬�[�찶]">����ק�ťլ��u�̬�[�찶]</a>
<hr>
<% else %>
 <form id="form1" name="form1" method="post" action="2_admin_main.asp" >
�n�J��x�޲z<br>
�K�X�G<input type="password" name="pwd" ><input type="submit" name="button1" value="�T�w">
</form>
<% end if %>

</center>
</BODY>
</HTML>