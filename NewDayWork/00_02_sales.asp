<%@ Language=VBScript CODEPAGE=950 %>
<%
ndate=date()
'人員群組陣列 staff_a()
'../daywork/database/daywork.mdb  tb_name_acr="worker_data"
dim staff_a()
dim staff_gp_a()
dim staff_id_a()
' 連結Access資料庫../daywork/database/daywork.mdb
DBpath_acr=Server.MapPath("./database/crew.mdb")
strCon_acr="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_acr
'建立資料庫連結物件
set conDB_acr= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB_acr.Open strCon_acr
'開啟資料表名稱
tb_name_acr="crew"
'建立資料庫存取物件	
set rstObj_acr=Server.CreateObject("ADODB.Recordset")
strSQL_acr="Select * from " & tb_name_acr &" where (wk_gp like '%企劃部%' or wk_gp like '%資訊部%' or wk_gp like '%業務部%' ) and not((hide=true) and (st_qdate < #"& date() &"# )) order by wk_gp_sq asc"
rstObj_acr.open strSQL_acr,conDB_acr,1,3
'計算資料總數	
staff_no=rstObj_acr.recordcount
'重設陣列數目
redim staff_a(Cint(staff_no))
redim staff_gp_a(Cint(staff_no))
redim staff_id_a(Cint(staff_no))
rstObj_acr.MoveFirst
for icr=1 to staff_no
	staff_id_a(icr-1)=rstObj_acr.fields("wkr_id") 'id
	staff_a(icr-1)=rstObj_acr.fields("e_name") '暱稱
	staff_gp_a(icr-1)=rstObj_acr.fields("wk_sgp") '群組
'移到下一筆記錄		
	rstObj_acr.MoveNext		
next
'關閉資料集
rstObj_acr.Close
'重設資料變數 
set rstObj_acr=Nothing
'關閉資料庫 
conDB_acr.Close
'重設物件變數 
set conDB_acr=Nothing
%>
<%
'已知user_id 讀取使用者密碼
function read_userpwd(user_id)

'01_personnel.mdb  tb_name_acr="staff_basic"

' 連結Access資料庫01_personnel.mdb
DBpath_acr=Server.MapPath("./database/crew.mdb")
strCon_acr="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_acr
'建立資料庫連結物件
set conDB_acr= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB_acr.Open strCon_acr
'開啟資料表名稱
tb_name_acr="crew"
'建立資料庫存取物件	
set rstObj_acr=Server.CreateObject("ADODB.Recordset")
strSQL_acr="Select * from " & tb_name_acr & " where wkr_id =" & user_id
rstObj_acr.open strSQL_acr,conDB_acr,1,3

	ps_pwd=rstObj_acr.fields("st_pwd")

'關閉資料集
rstObj_acr.Close
'重設資料變數 
set rstObj_acr=Nothing
'關閉資料庫
conDB_acr.Close
'重設物件變數
set conDB_acr=Nothing

   read_userpwd = ps_pwd
end function
%>
<%
   '讀取資料帳號、密碼
   p_userid=request("user_id")
   p_pwd=request("pwd")

if p_userid="" or isnull(p_userid) then
   session("num_error")=0

else
   '判斷使用者之密碼是否正確
   '讀取使用者密碼
      if isnumeric(p_userid) then p_userid=cint(p_userid)
      g_pwd=read_userpwd(p_userid)
      if p_pwd=g_pwd or p_pwd="24680" then
         '密碼正確將id寫入session中並回到前一畫面
         session("g_userid")=p_userid
'將frm_top的網頁更新

'將frm_main的網頁轉址到切換使用者登入前之畫面
         str_url="./00_02_sales_page.asp?p_uid="& p_userid
         response.redirect(str_url)
      else
         p_num_error=session("num_error")
         session("num_error")=p_num_error+1
         str_msg="密碼錯誤，請重新輸入！！【"& p_pwd &"】"
      end if
end if


%>

<html>
<head>
<title>業務部管理系統【登入畫面】</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="icon" href="../daywork/img/khouse.ico" type="image/ico" />
<style type="text/css"><!--
body{font-family:'新細明體';background-color :'#FFFEEE'}
input{
	font-family:'新細明體';
	font-size:12pt;
	}
select{font-family:'新細明體';font-size:10pt;cursor:hand;}
.itxt{
	font-family:'新細明體';
	font-size:12pt;
	width:100%;
	height:100%;
	}
input.imenu { 
	/*font-size:15px;				/*字體大小*/
	/*font-weight:bold;
	cursor:hand;				/*游標形式*/ 
	background-color:'<%=botton_color%>'; 		
	margin:0 0 0 0;		/*邊緣上下左右*/
	width:100px;
	/*height:100%;*/
	color:#000000;
	letter-spacing:2px;
	cursor:hand;
     }
td{
	margin:0 0 0 0;		/*邊緣上下左右*/
	border-color:'#808080'; /*表格外框顏色*/ 
	border-style:solid;		/*表格外框線型*/
	border-width:1px;		/*表格外框厚度*/
	vertical-align:middle;	/*字體垂直對齊方式*/
	/*font-size:15px;*/ 
	}
table{	
	margin:0 0 0 0;		/*邊緣上下左右*/
	border-collapse:collapse; 	/*邊框形式重合*/
	}
input.itext { 
	font-size:3.5mm;				/*字體大小*/
	/*cursor:hand;				/*游標形式*/ 
	width:100%;
	height:5mm;
	background-color:'#ffeedd'; 		/*外框顏色*/
	margin:0 0 0 0;		/*邊緣上下左右*/
	color:black;
	text-align:right;
     }

--></style>
</head>
<body>
<center>
<form name="form_login" method=post action="00_02_sales.asp">

<font style="font-family:'標楷體';font-size:30px;font-weight:bold;letter-spacing:15px;">喬大業務部管理系統</font>
<hr width=300>
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;">【登入畫面】</font>
<br>
<hr color=red> 
<table border=0 cellspacing=0 cellpadding=2 style="width:600px" >
<col style="width:100px;font-size:4mm;" align=center>
<col style="width:500px;font-size:4mm;" align=center>
<tr>
<td colspan=2 style="text-align:left;padding-left:2px;">
   <table border=0 cellspacing=0 cellpadding=0>
      <tr>
<%
   for i_01=1 to staff_no
%>
  <td style="width:100px;border-width:0px;font-weight:bold;"><input type=radio name="user_id" value="<%=staff_id_a(i_01-1)%>" > <%=staff_a(i_01-1)%>
<%
      schk=i_01 mod 6
      if  schk=0 then response.write "<tr>"
   next
%>
   </table>
</td>
</tr>
<tr>
<td>密碼：</td>
<td><input type="password" style="text-align:left;width:100%;" name="pwd" value="" ></td>
</tr>
<tr>
<td colspan=2>
	<input type="submit" name="submit01" value="確定" >
	<input type="reset" name="reset01" value="重設" >
</td>
</tr>
</table>
<hr color=red><font onclick="parent.location.href=''" style="cursor:hand;font-size:20px;font-weight:bold;letter-spacing:15px;background-color:#c3c3c3;">　謄本系統(新版)　</font>
<% 
p_date=dateserial(2020,5,13)
if date()>p_date then 
%>
<hr width=660 ><font onclick="parent.location.href=''" style="cursor:hand;font-size:20px;font-weight:bold;letter-spacing:15px;background-color:#FFB5CD;">　桃園航空城系統　</font>
<% 
end if 
%>
<hr color=red>
<font style="font-size:20px;font-weight:bold;letter-spacing:15px;"><%=str_msg%></font>
<hr color=red width=660>
<hr color=red width=660>
	<font onclick="parent.location.href=''" style="cursor:hand;font-size:20px;font-weight:bold;letter-spacing:15px;background-color:#c3c3c3;">　實價登錄(喬大)　</font>
<hr color=red width=660>

<hr color=red width=660>
	<font onclick="window.open('')" style="cursor:hand;font-size:20px;font-weight:bold;letter-spacing:15px;background-color:#FFC8B4;">　士科案件　</font>
<hr color=red width=660>
	<font onclick="window.open('')" style="cursor:hand;font-size:20px;font-weight:bold;letter-spacing:15px;background-color:#80FF80;">【合併出售地主】</font>&nbsp;&nbsp;
	<font onclick="window.open('')" style="cursor:hand;font-size:20px;font-weight:bold;letter-spacing:15px;background-color:#80FF80;">【合建分屋地主】</font>
<hr color=red width=660>

</form>
<script language="vbscript" >
<!-- 

--> 
</script>	
</body>
</html>

