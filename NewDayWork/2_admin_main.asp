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
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'標楷體';background-color:'#F0FFF0'}
--></style>
</HEAD>
<BODY>
<center>
<%
if session("admin_pwd")="28283939" then
%>
<a href="1_ulf_list.asp" >所有上傳檔案之列表</a><br>
<a href="1_ulf_del_list.asp" >已刪除之上傳檔案列表</a><br>
<a href="2_admin_main.asp?logout=1" title="登出後台管理">登出</a>
<hr>
<a href="wk_revise_order.asp" title="整體修改派工者名稱">整體修改派工者名稱</a>
<hr>
<a href="z_wk_lst_null.asp" title="派工者為空白之所有工作列表">派工者為空白之所有工作列表</a>
<hr>
<a href="z_wk_revise_order_null_ok.asp" title="整體修改空白派工者為[木偉]">整體修改空白派工者為[木偉]</a>
<hr>
<% else %>
 <form id="form1" name="form1" method="post" action="2_admin_main.asp" >
登入後台管理<br>
密碼：<input type="password" name="pwd" ><input type="submit" name="button1" value="確定">
</form>
<% end if %>

</center>
</BODY>
</HTML>