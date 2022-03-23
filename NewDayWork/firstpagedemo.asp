<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<!-- Include file = "./include/array_worker.inc" -->
<!-- Include file = "./include/array_worker_e.inc" -->

<%
	'設定Session變數消滅時間
	worker = Session("worker")
	today=date()
	tomorrow=date()+1
serverID=request.servervariables("LOCAL_ADDR")  

if serverID="192.168.0.139" then 
	serverID=""
else
	serverID="--<font color=red>【"&serverID&"】</font>"
end if
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
<link rel="icon" href="./img/khouse.ico" type="image/ico" />
<link rel="stylesheet" type="text/css" href="base_first.css" title="style1">
<style type="text/css"><!--
.ma1{
	font-family:'新細明體';
	color:red;
	font-size:24pt;
	letter-spacing:2mm;
	} 
.ma2{
	font-family:'新細明體';
	color:black;
	font-size:20pt;
	}
.ma3{
	font-family:'新細明體';
	color:black;
	font-size:10pt;
	}
.ma1a{
	font-family:'新細明體';
	color:red;
	font-size:24pt;
	letter-spacing:2mm;
	} 
.ma2a{
	font-family:'新細明體';
	color:blue;
	font-size:24pt;
	letter-spacing:2mm;
	}
.ma1z{
	font-family:'新細明體';
	color:red;
	font-size:12pt;
	} 
.ma2z{
	font-family:'新細明體';
	color:red;
	font-size:12pt;
	}
a:link    {color:blue;}
a:visited {color:blue;}
a:hover   {color:red;}
a:active  {color:green;}
--></style>
<script language="JavaScript">
</script>        
</HEAD>
<BODY topmargin=5>
	<FORM name="form1" action="work_main.asp" method=post >
<CENTER>
<!-- 
script language="JavaScript"

var popUpURL="./birthlist.asp";
splashWin =window.open("",'x','fullscreen=0,toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=auto,resizable=1');
splashWin.focus();
splashWin.location=popUpURL;
  
/script
 -->
<!-- 標誌圖片 -->
<img src = './img/work_tit.jpg' margin=0 style="">

<!-- 跑馬燈開始 -->
<!-- <marquee behavior="scroll" scrolldelay='210' BGCOLOR='#cff3c0' width=750> -->
<table border=1 cellspacing=0 cellpadding=0 width=783>
<% if headline_no=0 then %>
<tr><td>
<!--
		<marquee behavior="scroll" scrolldelay='210' BGCOLOR='#cff3c0' width=778 LOOP=3>
		<div>
		<font class='ma1'>喬大文化</font><font class='ma2'>：培養共同興趣、養成良好習慣。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
		<font class='ma1'>喬大精神</font><font class='ma2'>：做什麼、像什麼，用心做好每一件事情。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
		<font class='ma1'>喬大經營理念</font><font class='ma2'>：專業、負責、效率。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
		<font class='ma1'>喬大經營哲學</font><font class='ma2'>：利人、利他、利己。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
		<font class='ma1'>喬大致理格言</font><font class='ma2'>：忍耐是不夠的，必須學會包容。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
		<font class='ma1'>經營者的話</font><font class='ma2'>：</font><font class='ma3'>觀念</font><font class='ma2'> –是成敗的基本思維模式。</font><font class='ma3'>方法</font><font class='ma2'> – 是達成目標的重要選擇。</font><font class='ma3'>態度</font><font class='ma2'> – 是決定您命運的指南針。</font><font class='ma3'>誠信</font><font class='ma2'> – 是您我人格價值的象徵。</font><font class='ma3'>毅力</font><font class='ma2'> – 是成功關鍵的不二法門。</font><font class='ma3'>真心</font><font class='ma2'> – 是豐富人生快樂的泉源。</font>
		 </div>
		</marquee>
		-->
		<!--	
		<marquee behavior="scroll" scrolldelay='210' BGCOLOR='#cff3c0' width=778 LOOP=10>
		<div>
		<!--生日人員資料列表-->
		<!--		<font class='ma1'><%=pballname%></font>
		</div>
		</marquee>-->
<marquee behavior="scroll" scrolldelay='210' BGCOLOR='#cff3c0' width=778 LOOP=3>
<font class='ma1'>&nbsp;</font>
</marquee>
</td></tr>
<% else %>
<tr><td style="height:22pt;">
<!--
<marquee behavior="scroll" DIRECTION="up" scrolldelay='500' BGCOLOR='#cff3c0' width=778 LOOP="0" height="15">
 -->
<!-- <marquee behavior="scroll" DIRECTION="left" SCROLLAMOUNT="4" scrolldelay='150' BGCOLOR='#cff3c0' width=778 LOOP="10" height="22">-->
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
	<tr><td align=center><font class='tit' style="letter-spacing:10px;font-weight:bold;font-size:30px;font-family:'微軟正黑體';">喬大 <font style="letter-spacing:0px;"><font color=blue>G</font><font color=red>o</font><font color=yellow>o</font><font color=blue>g</font><font color=green>l</font><font color=red>e</font></font></font></td></tr>
	<tr>	
		<td align=center>
		<A Href='../Bulletin/80_main.asp?p_tpid=3' target=_blank style="background-color:#ccc;text-decoration:none;color:red;font-size:20px;letter-spacing:4px;">【佈告欄】</a>&nbsp;&nbsp;
<!--
 		<A Href='./birthlist-20081115a.asp' target=_blank>
			客戶今日及明日壽星列表
		</a>
-->
		<A Href='../customer/0_revise_birthday_list_now.asp' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			客戶今日及明日壽星列表</a>&nbsp;&nbsp;
		<A Href='../customer/0_revise_birthday_list.asp?chkmonth=<%=month(date())%>' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			[本月]</a>&nbsp;&nbsp;
		<A Href='../customer/0_revise_birthday_list.asp?chkmonth=<%=month(date())+1%>' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			[下月]</a>

		 </td>
	</tr>
	<tr><td align=center>
		<!-- #Include file = "./include/toolbar_worker_first_e.inc" -->
	</td></tr>
	</FORM>	
	<td align=center>
		<table>
		<tr align=center>
		<td width=150 style="background-color:#DDDDDD;"><A Href="./分機表.asp" target="_blank" title="公司人員分機及電子郵件資料表" style="color:#000000;text-decoration:none;">分機表</a></td>
		<td width=150 style="background-color:#FFCCCC;"><A Href='./00_01_建設.asp' target=_blank style="color:#000000;text-decoration:none;">建設部</A></td>
		<td width=150 style="background-color:#FFDDAA;"><A Href='./00_02_業務demo.asp' target=_blank style="color:#000000;text-decoration:none;font-family:微軟正黑體;font-weight:bold;">冒泡部</A></td>
		<td width=150 style="background-color:#FFFFBB;"><A Href='./00_03_管理.asp' target=_blank style="color:#000000;text-decoration:none;">管理部</A></td>
		<td width=150 style="background-color:#CCFF99;"><A Href='./00_04_財務.asp' target=_blank style="color:#000000;text-decoration:none;">財務部</A></td>
		<td width=150 style="background-color:#BBFFEE;"><A Href='./00_05_法務.asp' target=_blank style="color:#000000;text-decoration:none;">法務部</A></td>
		<td width=150 style="background-color:#99FFFF;"><A Href='./00_06_資訊.asp' target=_blank style="color:#000000;text-decoration:none;">資訊部</A></td>
		<td width=150 style="background-color:#CCCCFF;"><A Href='./00_07_高爾夫.asp' target=_blank style="color:#000000;text-decoration:none;">高爾夫</A></td>
		<td width=150 style="background-color:#FFB3FF;"><A Href='./00_08_基金會.asp' target=_blank style="color:#000000;text-decoration:none;">基金會</A></td>
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
	<img src="./img/link.png" style="height:25px;vertical-align:middle;" onclick="sh01.style.display=sh01.style.display=='none'?'':'none'" title="顯示其他連結">
</TD>
</TR></TABLE>
<div id="sh01" style="display:none;padding-left:0px;">
<!-- #Include file = "./連結網頁.asp" -->
</div>


</center>

</BODY>
</HTML>






