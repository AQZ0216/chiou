<% @codepage=950%>
<!-- #Include file = "./include/array_worker.inc" -->
<!-- #Include file = "./include/array_worker_e.inc" -->
<%
	'�]�wSession�ܼƮ����ɶ�
	worker = Session("worker")
%>

<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<link rel="stylesheet" type="text/css" 
href="base_first.css" title="style1">
<style type="text/css"><!--
.ma1{
	font-family:'�s�ө���';
	color:red;
	font-size:12pt;
	} 
.ma2{
	font-family:'�s�ө���';
	color:black;
	font-size:10pt;
	} 
.ma3{
	font-family:'�s�ө���';
	color:black;
	font-size:10pt;
	} 
--></style>
</HEAD>
<BODY>

<CENTER>
<!-- �лx�Ϥ� -->
<CENTER>
<img src = './img/work_tit.jpg'>
</CENTER>
	<FORM name="form1" action="work_main.asp" method=post >
<TABLE BORDER=1 cellspacing=0 cellpadding=0>
<col width=600><col width=150>
<TR><TD VALIGN=TOP>

	<center>
	<table border=1>
	<col width=600>
	<tr><td align=center><font size=5>�w��i�J�ӤH�u�@�޲z�t��!!</font></td></tr>
	<tr>
		<td align=center><% =worker %> �A�n�I</td>
	</tr>
	<tr><td align=center>
		<!-- #Include file = "./include/toolbar_worker_first_e.inc" -->
	</td></tr>
	</FORM>	
	<td align=center>
		<table>
		<tr align=center>
		<td width=150><A Href='wkr_add.asp' target=_top>�u�@�H���ק�</A></td>
		<td width=150><A Href='../customer/login.asp' target=_top>�Ȥ��ƽs��</A></td>
		<td width=150><A Href="file://Chiou-server/work_control/�g�a���.xls" target="_blank">�D���Φa��J</a></td>
		<td width=150><A Href='../tel_message/login_msg.asp' target=_blank>����Ȥ�ӹq</A></td>
		</tr>
		</table>
	</td>
	</table>
    	</center>
    	
    </TD>
<TD align=center>
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
	     document.write("<TABLE CELLSPACING=3 CELLPADDING=2>");
	     document.write("<TR ",thcol,"><TH COLSPAN=7><font size=1>","����"+ROCYear+"�~",names[Month],thisDay,"��","</font></th>");
	     document.write("<TR ",trcol,"><TH><font size=1>��</font></TH><TH><font size=1>�@</font></TH><TH><font size=1>�G</font></TH><TH><font size=1>�T</font></TH><TH><font size=1>�|</font></TH><TH><font size=1>��</font></TH><TH><font size=1>��</font></TH></TR>");
	     document.write("<TR ALIGN=RIGHT>");
	     var column = 0;
	     for (i=0; i<startDay; i++)
	     {
	          document.write("<TD><font size=1>&nbsp</font></TD>");
	          column++;
	     }
	     for (i=1; i<=days[Month]; i++)
	     {
	          if ((i == thisDay)  && (Month == thisMonth))
	               document.write("<TD ",tocol,"><font size=1>",i,"</font></TD>");
	          else
	             {
	               if ((column == 0) || (column == 6))
	                 document.write("<TD ",hlcol,"><font size=1>",i,"</font></TD>");
	               else
	                 document.write("<TD ",tdcol,"><font size=1>",i,"</font></TD>");
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
	
	var names = new array("1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��");
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
	if (thisYear < 2000) thisYear=thisYear+13;
	Calendar(thisMonth,thisYear);
	
	//-->
	</SCRIPT>
</TD>
</TR></TABLE>
</CENTER>
<center>
<table border="1" cellspacing=1 cellpadding=1>
<col width=150><col width=150><col width=150><col width=150><col width=150>
	<tr align=center>
	<A Href='http://tw.yahoo.com' target=_blank><td class=urlcmd>Yahoo�_��!!</td></A>
	<A Href='http://google.com' target=_blank><td class=urlcmd>google</td></A>
	<A Href='http://www.pchome.com.tw' target=_blank><td class=urlcmd>pchome</td></A>
	<A Href='http://www.khouse.com.tw' target=_blank><td class=urlcmd>��j����</td></A>
	<A Href='http://lhouse.com.tw' target=_blank><td class=urlcmd>�e�W�k��</td></A>	
	</tr>
	<tr align=center>
	<A Href='http://www.landagent.com.tw/vip-room/la-vip.asp' target=_blank><td class=urlcmd>�{�N�a�F</td></A>
	<A Href='http://www.zone.taipei.gov.tw/' target=_blank><td class=urlcmd>�ϥΤ���</td></A>
	<A Href='http://www.land.taipei.gov.tw/tgl00000.asp?page=d&no=1' target=_blank><td class=urlcmd>���i�{�Ȭd��</td></A>	
	<A Href='http://land.hinet.net' target=_blank><td class=urlcmd>���٦a�y�å�</td></A>	
	<A Href='http://www.tsland.gov.tw/f/f.htm' target=_blank><td class=urlcmd>���٦a�F�B</td></A>
	</tr>
	<tr align=center>
	<A Href='http://www.taipei.gov.tw/' target=_blank><td class=urlcmd>�x�_���F��</td></A>
	<A Href='http://www.land.taipei.gov.tw/tgl00000.asp?page=d&no=1' target=_blank><td class=urlcmd>�x�_���a�F�B</td></A>
	<A Href='http://163.29.37.132/html/main.htm' target=_blank><td class=urlcmd>�x�_���غ޳B</td></A>	
	<A Href='http://www.tpctax.gov.tw/index.htm' target=_blank><td class=urlcmd>�_���|���]�x�B</td></A>
	<A Href='http://www.planning.taipei.gov.tw/TCDB_C/default.asp' target=_blank><td class=urlcmd>�_�������o�i��</td></A>

	</tr>
	<tr align=center>
	<A Href='http://www.tsland.gov.tw/e/e1.htm' target=_blank><td class=urlcmd>�g�a�W�ȵ|�պ�</td></A>
	<A Href='http://www.houseno.tcg.gov.tw' target=_blank><td class=urlcmd>���P�˯��t��</td></A>	
	<A Href='http://www.ntat.gov.tw/' target=_blank><td class=urlcmd>�x�_����|��</td></A>
	<A Href='http://egw20.mofdpc.gov.tw/bgq/' target=_blank><td class=urlcmd>��~�n�O���</td></A>
	<A Href='http://law.moj.gov.tw/fp.asp' target=_blank><td class=urlcmd>����k�W��Ʈw</td></A>	
	</tr>
	<tr align=center>
	<A Href='http://www.consumers.org.tw/' target=_blank><td class=urlcmd>����|</td></A>
	<A Href='http://www.judicial.gov.tw/' target=_blank><td class=urlcmd>�q�k�|</td></A>
	<A Href='http://www.ftc.gov.tw/' target=_blank><td class=urlcmd>��������e���|</td></A>
	<A Href='./qa/qa_all.html' target=_blank><td class=urlcmd>�q���ާ@���D��</td></A>
	<A Href='./learn/learn_web.html' target=_blank><td class=urlcmd>�q���оǺ���</td></A>
	</tr>

</table>
</center>

</BODY>
</HTML>
