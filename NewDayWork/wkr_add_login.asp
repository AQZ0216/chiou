<% @codepage=950%>
<%
pwd=request("pwd01")
if pwd="6980" then
   str_url="wkr_add_new.asp" '�i�J�s�W�H����Ƶe��
else
   errmsg="�K�X���~!!�п�J���T�K�X!!"
   str_url="wkr_login.asp?errmsg="&errmsg '�i�J�s�W�H����Ƶe��
end if
   response.redirect(str_url)

%>