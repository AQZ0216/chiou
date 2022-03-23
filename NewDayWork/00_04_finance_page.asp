<%@ Language=VBScript CODEPAGE=950 %>
<%
p_uid=request("p_uid")
if p_uid="" then
	str_url="./00_04_finance.asp"
  response.redirect(str_url)
end if

%>
<%
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
strSQL_acr="Select * from " & tb_name_acr &" where wkr_id = " & p_uid
rstObj_acr.open strSQL_acr,conDB_acr,1,1
'計算資料總數	
staff_no=rstObj_acr.recordcount
rstObj_acr.MoveFirst
for icr=1 to staff_no
	staff_name=rstObj_acr.fields("wkr_name") '暱稱
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

<HTML>
<HEAD>
<title>財務部系統首頁</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" 
href="base_first.css" title="style1">
<style type="text/css"><!--
.tde{
	font-family:'標楷體';
	/*color:red;*/
	/*background-color:#DDDDDD;"*/
	font-size:15pt;
	} 
a.tda{
	text-decoration:none; 
	}
a:link    {color:blue;}
a:visited {color:blue;}
a:hover   {color:red;}
a:active  {color:green;}
--></style>
<script language="JavaScript">
</script>        
</HEAD>
<BODY topmargin=15>
<CENTER>
<TABLE BORDER=1 cellspacing=0 cellpadding=0 width=783>
			<tr style="height:35pt;">
				<td align=center style="background-color:#CCFF99;"><font style="font-size:20pt;letter-spacing:2mm;">歡迎進入財務部管理系統!!</font></td>
			</tr>

</table>
<TABLE BORDER=1 cellspacing=0 cellpadding=0 width=783>
<col width=600><col width=160>
<TR>
	<TD VALIGN=TOP style="text-align:center;">
		<table border=1>
		<col width=590>
			<tr>
				<td align=center>
				<table border=1 style="width:100%">
						<col style="width:25%;">
						<col style="width:25%;">
						<col style="width:25%;">
						<col style="width:25%;">
					<tr align=center style="height:25pt;" >
						<td class=tde ><A class=tda Href='' target=_blank>業績資料</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >銷售案件</A></td>
						<td class=tde ><A class=tda Href='' target=_blank>代書案件</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >代書業務記錄簿</A></td>
					</tr>
					<tr align=center style="height:25pt;" >

						<td class=tde ><A class=tda Href='' target='_blank' style="color:red;background-color:#e1e1e1;" >【業績日期】</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' style="color:red;background-color:#e1e1e1;" > 【案件稽核】</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
					</tr>
					<tr align=center style="height:25pt;" >
						<td class=tde ><A class=tda Href='../expenses/a01_main.asp' target='_blank' style="color:red;background-color:#e1e1e1;" > 【支出證明單】</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
					</tr>
					<tr align=center style="height:25pt;" >
						<td class=tde ><A class=tda Href='../rights/a00_login.asp' target='_blank' style="color:red;background-color:#e1e1e1;" > 【權狀管理】</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
					</tr>
					<tr align=center style="height:25pt;" >
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
					</tr>
					<tr align=center style="height:25pt;" >
						<td class=tde ><A class=tda Href='' target='_blank' title="喬大票據管理系統">喬大票據管理</A></td>
						<td class=tde ><A class=tda Href='' target=_top>客戶資料</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >基本範本檔</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >回首頁</A></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
  </TD>
	<TD align=center valign=middle>
		<SCRIPT LANGUAGE="JavaScript"><!--
		
		function Calendar(Month,Year)
		{
		     if (Year < 1900)
		         Year=Year+1900;
		     firstDay = new Date(Year,Month,1);
		     startDay = firstDay.getDay();
		     if (((Year % 4 == 0) && (Year % 100 != 0)) || (Year % 400 == 0))
		          days[1] = 29; 
		     else
		          days[1] = 28;
		     ROCYear=Year-1911;
		     document.write("<TABLE CELLSPACING=3 CELLPADDING=2 >");
		     document.write("<TR ",thcol,"><TH COLSPAN=7><font style='font-size:12pt;'>","民國"+ROCYear+"年",names[Month],thisDay,"日","</font></th>");
		     document.write("<TR ",trcol,"><TH><font style='font-size:11pt;'>日</font></TH><TH><font style='font-size:11pt;'>一</font></TH><TH><font style='font-size:11pt;'>二</font></TH><TH><font style='font-size:11pt;'style='font-size:11pt;'>三</font></TH><TH><font style='font-size:11pt;'>四</font></TH><TH><font style='font-size:11pt;'>五</font></TH><TH><font style='font-size:11pt;'>六</font></TH></TR>");
		     document.write("<TR ALIGN=RIGHT>");
		     var column = 0;
		     for (i=0; i<startDay; i++)
		     {
		          document.write("<TD><font style='font-size:11pt;'>&nbsp</font></TD>");
		          column++;
		     }
		     for (i=1; i<=days[Month]; i++)
		     {
		          if ((i == thisDay)  && (Month == thisMonth))
		               document.write("<TD ",tocol,"><font style='font-size:11pt;'>",i,"</font></TD>");
		          else
		             {
		               if ((column == 0) || (column == 6))
		                 document.write("<TD ",hlcol,"><font sizstyle='font-size:11pt;'>",i,"</font></TD>");
		               else
		                 document.write("<TD ",tdcol,"><font style='font-size:11pt;'>",i,"</font></TD>");
		             }
		          column++;
		          if (column == 7)
		          {
		               document.write("</TR><TR ALIGN=RIGHT>");
		               column = 0;
		          }
		     }
		     document.write("</TR></TABLE>");
		}
		
		function array(m0, m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11)
		{
		     this[0] = m0; this[1] = m1; this[2]  = m2;  this[3]  = m3;
		     this[4] = m4; this[5] = m5; this[6]  = m6;  this[7]  = m7;
		     this[8] = m8; this[9] = m9; this[10] = m10; this[11] = m11;
		}
		
		var names = new array("1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月");
		var days  = new array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
		var thcol = "BGCOLOR='#ffc080'";
		var trcol = "BGCOLOR='#c0ffc0'";
		var tdcol = "BGCOLOR='#ffffff'";
		var tocol = "BGCOLOR='#ffc080'";
		var hlcol = "BGCOLOR='#ffc0c0'";
		var today     = new Date();
		var thisDay   = today.getDate();
		var thisMonth = today.getMonth();
		var thisYear  = today.getYear() ;
		//if (thisYear < 2000) thisYear=thisYear+13;
		Calendar(thisMonth,thisYear);
		
		//-->
		</SCRIPT>
	</TD>
</TR>
</TABLE>

</center>

</BODY>
</HTML>






