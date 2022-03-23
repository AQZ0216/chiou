<% @codepage=950%>
<%
pwd=request("pwd01")
if pwd="6980" then
   str_url="wkr_add_new.asp" '進入新增人員資料畫面
else
   errmsg="密碼錯誤!!請輸入正確密碼!!"
   str_url="wkr_login.asp?errmsg="&errmsg '進入新增人員資料畫面
end if
   response.redirect(str_url)

%>