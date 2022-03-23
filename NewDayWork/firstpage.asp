<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->

<%
	'設定Session變數消滅時間
	worker = Session("worker")
	today=date()
	tomorrow=date()+1
serverID=request.servervariables("LOCAL_ADDR")  


userip = Request.ServerVariables("REMOTE_ADDR") '本機ip
%>
<%
'headline_no   '重大訊息數量
dim headline_txt()      '重大訊息內容
dim headline_date()   '重大訊息日期

      ' 連結Access資料庫daywork.mdb
      DBpath_hdl=Server.MapPath("./database/daywork.mdb")
      strCon_hdl="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_hdl
      '建立資料庫連結物件
      set conDB_hdl= Server.CreateObject("ADODB.Connection")
      '連結資料庫	
      conDB_hdl.Open strCon_hdl
      '開啟資料表名稱
      tb_name_hdl="work_data"
      '建立資料庫存取物件	
      set rstObj1_hdl=Server.CreateObject("ADODB.Recordset")
'      strSQL_show_hdl="Select * from " & tb_name_hdl & " where headline > 5  order by doing_date1 asc"
      strSQL_show_hdl="Select * from " & tb_name_hdl & " where headline > 5 and doing_date1 = #"& date() &"# order by wk_item asc"   
      rstObj1_hdl.open strSQL_show_hdl,conDB_hdl,1,3
totalput_hdl=rstObj1_hdl.recordcount
headline_no=totalput_hdl
'response.write "strSQL_show_hdl="& strSQL_show_hdl &"。" 
'response.write "headline_no="& headline_no &"。" 

if totalput_hdl=0 then
   redim headline_txt(1)
   redim headline_date(1)
   headline_date(0)=date()
   headline_txt(0)="無"
else
   redim headline_txt(headline_no)
   redim headline_date(headline_no)
	'列出資料項目
	rstobj1_hdl.MoveFirst
	for i=1 to totalput_hdl
	     headline_date(i-1)=rstObj1_hdl.fields("doing_date1")
	 	headline_txt(i-1)=rstObj1_hdl.fields("wk_item")
	'移到下一筆記錄
		rstObj1_hdl.MoveNext
		if rstObj1_hdl.EOF=True then exit for
	next	
end if	
      '關閉資料集
      rstObj1_hdl.Close
      '重設資料變數
      set rstObj1_hdl=Nothing
      '關閉資料庫
      conDB_hdl.Close
      '重設物件變數
      set conDB_hdl=Nothing

%>

<%
if userip="127.0.0.2" then
	pballname=""		
else
'--'列出今日生日壽星之姓名字串 pballname
	'列出今日生日壽星之姓名字串 pballname
	'客戶資料庫customer_data.mdb 資料表customer_tb
	' 連結Access資料庫./database/customer_data.mdb
'	DBpath=Server.MapPath("../customer/database/customer_data.mdb")
'	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'建立資料庫連結物件
'	set conDB= Server.CreateObject("ADODB.Connection")
	'連結資料庫	
'	conDB.Open strCon
	'開啟資料表名稱
'	tb_name="customer_tb"
	'建立資料庫存取物件	
'	set rstObj1=Server.CreateObject("ADODB.Recordset")
'	str_da0=" (month(c_birthday_dt)=month(date()) and day(c_birthday_dt)=day(date())) "
'	strSQL_list="Select * from " & tb_name & " where "& str_da0&" order by c_name asc"
'	rstObj1.open strSQL_list,conDB,3,1
	'計算資料總數	
'	totalput=rstObj1.recordcount
'		pballname=""
'	if totalput= 0 then
'	else
		'移至第一筆資料
'		rstObj1.MoveFirst
'	    for kj=1 to totalput
'	    	pname="　祝　<font color=blue>" & rstObj1.fields("c_name")& "</font>　生日快樂!!　"         '
'	    	pballname=pballname & "　" & pname
	      '移到下一筆記錄
'	      rstObj1.MoveNext
'	      if rstObj1.EOF=True then exit for
'	    next
'	end if
	
	'關閉資料集
'	rstObj1.Close
	'重設資料變數 
'	set rstObj1=Nothing
	'關閉資料庫 
'	conDB.Close
	'重設物件變數 
'	set conDB=Nothing
'--'列出今日來訪資料 字串 pballname
	'列出今日來訪資料 字串 pballname
	'daywork.mdb 資料表work_data
	' 連結Access資料庫./database/daywork.mdb
	DBpath=Server.MapPath("./database/daywork.mdb")
	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'建立資料庫連結物件
	set conDB= Server.CreateObject("ADODB.Connection")
	'連結資料庫	
	conDB.Open strCon
	'開啟資料表名稱
	tb_name="work_data"
	'建立資料庫存取物件	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	str_da0=" doing_date1 = date() "
	strSQL_list="Select * from " & tb_name & " where "& str_da0 &" order by wk_item asc"
	rstObj1.open strSQL_list,conDB,3,1
	'計算資料總數
	totalput=rstObj1.recordcount
		pballname=""
	if totalput= 0 then
	else
			pballname= "來訪訊息："
		'移至第一筆資料
		rstObj1.MoveFirst
	    for kj=1 to totalput
	    	pp_kitem=rstObj1.fields("wk_item")
	    	if ( instr(1,pp_kitem,"來訪",1)>0 or instr(1,pp_kitem,"到公司",1)>0 ) then 	
		    	pname="<font color=blue>" & rstObj1.fields("wk_item")& "</font>"         '
		    	pballname=pballname & "　" & pname
	      end if
	      '移到下一筆記錄
	      rstObj1.MoveNext
	      if rstObj1.EOF=True then exit for
	    next
	end if
	
	'關閉資料集
	rstObj1.Close
	'重設資料變數 
	set rstObj1=Nothing
	'關閉資料庫 
	conDB.Close
	'重設物件變數 
	set conDB=Nothing

end if
%>

<HTML>
<HEAD>
<title>工作管理系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="icon" href="../daywork/img/khouse.ico" type="image/ico" />
<link rel="stylesheet" type="text/css" href="./css/base_first.css" title="style1">
<style type="text/css"><!--
.ma1{
	font-family:'微軟正黑體';
	color:red;
	font-size:24pt;
	letter-spacing:2mm;
	} 
.ma2{
	font-family:'微軟正黑體';
	color:black;
	font-size:20pt;
	}
.ma3{
	font-family:'微軟正黑體';
	color:black;
	font-size:10pt;
	}
.ma1a{
	font-family:'微軟正黑體';
	color:red;
	font-size:24pt;
	letter-spacing:2mm;
	} 
.ma2a{
	font-family:'微軟正黑體';
	color:blue;
	font-size:24pt;
	letter-spacing:2mm;
	}
.ma1z{
	font-family:'微軟正黑體';
	color:red;
	font-size:12pt;
	} 
.ma2z{
	font-family:'微軟正黑體';
	color:red;
	font-size:12pt;
	}
a:link    {color:blue;}
a:visited {color:blue;}
a:hover   {color:red;}
a:active  {color:green;}
a{text-decoration:none;}
--></style>
<script language="JavaScript">
</script>        
</HEAD>
<BODY topmargin=5>
	<FORM name="form1" action="work_main.asp" method=post >
<CENTER>

<!-- 標誌圖片 -->
<img src = './img/work_tit.jpg' margin=0 style="">

<!-- 跑馬燈開始 -->

<table border=1 cellspacing=0 cellpadding=0 width=783>
<% if headline_no=0 then %>
<tr><td>
<marquee behavior="scroll" scrolldelay='210' BGCOLOR='#cff3c0' width=778 LOOP=3>
<font class='ma1'>&nbsp;</font>
</marquee>
</td></tr>
<% else %>
<tr><td style="height:22pt;">
		<marquee behavior="scroll" scrolldelay="120" BGCOLOR="#cff3c0" width="778" LOOP="10">
<div>
<font class='ma1a'>訊息公告(<%=totalput_hdl%>筆)：</font><font class='ma2a'>&nbsp;&nbsp;</font>
<% for zi=1 to headline_no %>
<font class='ma1a'><%=zi%><!--<%=headline_date(zi-1)%>--></font>、<font class='ma2a'><%=headline_txt(zi-1)%>。&nbsp;&nbsp;</font>
<% next %>
</div>
</marquee>
</td></tr>
<% end if %>
</table>
<!-- 跑馬燈結束 -->


<TABLE BORDER=1 cellspacing=0 cellpadding=0 width=783>
<col width=600><col width=160>
<TR><TD VALIGN=TOP>

	<center>
	<table border=1>
	<col width=590>
	<tr><td align=center><font class='tit'>歡迎進入個人工作管理系統!!</font></td></tr>
	<tr>	
		<td align=center>
		<A Href='' target=_blank style="background-color:#ccc;text-decoration:none;color:red;font-size:20px;letter-spacing:4px;">【佈告欄】</a>&nbsp;&nbsp;

		<A Href='' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			客戶今日及明日壽星列表</a>&nbsp;&nbsp;
		<A Href='' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			[本月]</a>&nbsp;&nbsp;
		<A Href='' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			[下月]</a>
		 </td>
	</tr>
	<tr><td align=center>
		<!-- #Include file = "./include/toolbar_worker_first_e.inc" -->
	</td></tr>
	</FORM>	
	<td align=center>
		<table style="">
		<tr align=center>
		<td width=150 style="background-color:#FF79FF;"><A Href="" target="_blank" title="用印申請" style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">用印</a></td>
		<td width=150 style="background-color:#FFCCCC;"><A Href='./00_01_build.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">建設部</A></td>
		<td width=150 style="background-color:#FFDDAA;"><A Href='./00_02_sales.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">業務部</A></td>
		<td width=150 style="background-color:#FFFFBB;"><A Href='./00_03_manager.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">管理部</A></td>
		<td width=150 style="background-color:#CCFF99;"><A Href='./00_04_finance.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">財務部</A></td>
		<td width=150 style="background-color:#BBFFEE;"><A Href='./00_05_law.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">法務部</A></td>
		<td width=150 style="background-color:#99FFFF;"><A Href='./00_06_mis.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">資訊部</A></td>
		<td width=150 style="background-color:#CCCCFF;"><A Href='./00_07_golf.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">高爾夫</A></td>
		<td width=150 style="background-color:#FFB3FF;"><A Href='./00_08_fundation.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'微軟正黑體';font-weight:bold;">社 企</A></td>
		</tr>
		</table>
	</td>
	</table>
    	</center>
    	
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
	<img src="./img/demo.png" style="height:25px;vertical-align:middle;" onclick="demo()" title="登入展示版網頁">

</TD>
</TR></TABLE>

<!-- #Include file = "./連結網頁.asp" -->



</center>
<script language=vbscript>
sub demo() '展示板
	MyVar = MsgBox ("確定進入展示版本網頁！！。", 64+1, "MsgBox Example")
		   if MyVar =1 then
		   	'確定進入展示版本網頁
		   	location.href=""
		   end if	
end sub
</script>
</BODY>
</HTML>






