<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->

<%
	'�]�wSession�ܼƮ����ɶ�
	worker = Session("worker")
	today=date()
	tomorrow=date()+1
serverID=request.servervariables("LOCAL_ADDR")  


userip = Request.ServerVariables("REMOTE_ADDR") '����ip
%>
<%
'headline_no   '���j�T���ƶq
dim headline_txt()      '���j�T�����e
dim headline_date()   '���j�T�����

      ' �s��Access��Ʈwdaywork.mdb
      DBpath_hdl=Server.MapPath("./database/daywork.mdb")
      strCon_hdl="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_hdl
      '�إ߸�Ʈw�s������
      set conDB_hdl= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB_hdl.Open strCon_hdl
      '�}�Ҹ�ƪ�W��
      tb_name_hdl="work_data"
      '�إ߸�Ʈw�s������	
      set rstObj1_hdl=Server.CreateObject("ADODB.Recordset")
'      strSQL_show_hdl="Select * from " & tb_name_hdl & " where headline > 5  order by doing_date1 asc"
      strSQL_show_hdl="Select * from " & tb_name_hdl & " where headline > 5 and doing_date1 = #"& date() &"# order by wk_item asc"   
      rstObj1_hdl.open strSQL_show_hdl,conDB_hdl,1,3
totalput_hdl=rstObj1_hdl.recordcount
headline_no=totalput_hdl
'response.write "strSQL_show_hdl="& strSQL_show_hdl &"�C" 
'response.write "headline_no="& headline_no &"�C" 

if totalput_hdl=0 then
   redim headline_txt(1)
   redim headline_date(1)
   headline_date(0)=date()
   headline_txt(0)="�L"
else
   redim headline_txt(headline_no)
   redim headline_date(headline_no)
	'�C�X��ƶ���
	rstobj1_hdl.MoveFirst
	for i=1 to totalput_hdl
	     headline_date(i-1)=rstObj1_hdl.fields("doing_date1")
	 	headline_txt(i-1)=rstObj1_hdl.fields("wk_item")
	'����U�@���O��
		rstObj1_hdl.MoveNext
		if rstObj1_hdl.EOF=True then exit for
	next	
end if	
      '������ƶ�
      rstObj1_hdl.Close
      '���]����ܼ�
      set rstObj1_hdl=Nothing
      '������Ʈw
      conDB_hdl.Close
      '���]�����ܼ�
      set conDB_hdl=Nothing

%>

<%
if userip="127.0.0.2" then
	pballname=""		
else
'--'�C�X����ͤ�جP���m�W�r�� pballname
	'�C�X����ͤ�جP���m�W�r�� pballname
	'�Ȥ��Ʈwcustomer_data.mdb ��ƪ�customer_tb
	' �s��Access��Ʈw./database/customer_data.mdb
'	DBpath=Server.MapPath("../customer/database/customer_data.mdb")
'	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'�إ߸�Ʈw�s������
'	set conDB= Server.CreateObject("ADODB.Connection")
	'�s����Ʈw	
'	conDB.Open strCon
	'�}�Ҹ�ƪ�W��
'	tb_name="customer_tb"
	'�إ߸�Ʈw�s������	
'	set rstObj1=Server.CreateObject("ADODB.Recordset")
'	str_da0=" (month(c_birthday_dt)=month(date()) and day(c_birthday_dt)=day(date())) "
'	strSQL_list="Select * from " & tb_name & " where "& str_da0&" order by c_name asc"
'	rstObj1.open strSQL_list,conDB,3,1
	'�p�����`��	
'	totalput=rstObj1.recordcount
'		pballname=""
'	if totalput= 0 then
'	else
		'���ܲĤ@�����
'		rstObj1.MoveFirst
'	    for kj=1 to totalput
'	    	pname="�@���@<font color=blue>" & rstObj1.fields("c_name")& "</font>�@�ͤ�ּ�!!�@"         '
'	    	pballname=pballname & "�@" & pname
	      '����U�@���O��
'	      rstObj1.MoveNext
'	      if rstObj1.EOF=True then exit for
'	    next
'	end if
	
	'������ƶ�
'	rstObj1.Close
	'���]����ܼ� 
'	set rstObj1=Nothing
	'������Ʈw 
'	conDB.Close
	'���]�����ܼ� 
'	set conDB=Nothing
'--'�C�X����ӳX��� �r�� pballname
	'�C�X����ӳX��� �r�� pballname
	'daywork.mdb ��ƪ�work_data
	' �s��Access��Ʈw./database/daywork.mdb
	DBpath=Server.MapPath("./database/daywork.mdb")
	strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
	'�إ߸�Ʈw�s������
	set conDB= Server.CreateObject("ADODB.Connection")
	'�s����Ʈw	
	conDB.Open strCon
	'�}�Ҹ�ƪ�W��
	tb_name="work_data"
	'�إ߸�Ʈw�s������	
	set rstObj1=Server.CreateObject("ADODB.Recordset")
	str_da0=" doing_date1 = date() "
	strSQL_list="Select * from " & tb_name & " where "& str_da0 &" order by wk_item asc"
	rstObj1.open strSQL_list,conDB,3,1
	'�p�����`��
	totalput=rstObj1.recordcount
		pballname=""
	if totalput= 0 then
	else
			pballname= "�ӳX�T���G"
		'���ܲĤ@�����
		rstObj1.MoveFirst
	    for kj=1 to totalput
	    	pp_kitem=rstObj1.fields("wk_item")
	    	if ( instr(1,pp_kitem,"�ӳX",1)>0 or instr(1,pp_kitem,"�줽�q",1)>0 ) then 	
		    	pname="<font color=blue>" & rstObj1.fields("wk_item")& "</font>"         '
		    	pballname=pballname & "�@" & pname
	      end if
	      '����U�@���O��
	      rstObj1.MoveNext
	      if rstObj1.EOF=True then exit for
	    next
	end if
	
	'������ƶ�
	rstObj1.Close
	'���]����ܼ� 
	set rstObj1=Nothing
	'������Ʈw 
	conDB.Close
	'���]�����ܼ� 
	set conDB=Nothing

end if
%>

<HTML>
<HEAD>
<title>�u�@�޲z�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="icon" href="../daywork/img/khouse.ico" type="image/ico" />
<link rel="stylesheet" type="text/css" href="./css/base_first.css" title="style1">
<style type="text/css"><!--
.ma1{
	font-family:'�L�n������';
	color:red;
	font-size:24pt;
	letter-spacing:2mm;
	} 
.ma2{
	font-family:'�L�n������';
	color:black;
	font-size:20pt;
	}
.ma3{
	font-family:'�L�n������';
	color:black;
	font-size:10pt;
	}
.ma1a{
	font-family:'�L�n������';
	color:red;
	font-size:24pt;
	letter-spacing:2mm;
	} 
.ma2a{
	font-family:'�L�n������';
	color:blue;
	font-size:24pt;
	letter-spacing:2mm;
	}
.ma1z{
	font-family:'�L�n������';
	color:red;
	font-size:12pt;
	} 
.ma2z{
	font-family:'�L�n������';
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

<!-- �лx�Ϥ� -->
<img src = './img/work_tit.jpg' margin=0 style="">

<!-- �]���O�}�l -->

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
<font class='ma1a'>�T�����i(<%=totalput_hdl%>��)�G</font><font class='ma2a'>&nbsp;&nbsp;</font>
<% for zi=1 to headline_no %>
<font class='ma1a'><%=zi%><!--<%=headline_date(zi-1)%>--></font>�B<font class='ma2a'><%=headline_txt(zi-1)%>�C&nbsp;&nbsp;</font>
<% next %>
</div>
</marquee>
</td></tr>
<% end if %>
</table>
<!-- �]���O���� -->


<TABLE BORDER=1 cellspacing=0 cellpadding=0 width=783>
<col width=600><col width=160>
<TR><TD VALIGN=TOP>

	<center>
	<table border=1>
	<col width=590>
	<tr><td align=center><font class='tit'>�w��i�J�ӤH�u�@�޲z�t��!!</font></td></tr>
	<tr>	
		<td align=center>
		<A Href='' target=_blank style="background-color:#ccc;text-decoration:none;color:red;font-size:20px;letter-spacing:4px;">�i�G�i��j</a>&nbsp;&nbsp;

		<A Href='' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			�Ȥᤵ��Ω���جP�C��</a>&nbsp;&nbsp;
		<A Href='' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			[����]</a>&nbsp;&nbsp;
		<A Href='' target=_blank style="background-color:#ddd;text-decoration:none;color:blue;">
			[�U��]</a>
		 </td>
	</tr>
	<tr><td align=center>
		<!-- #Include file = "./include/toolbar_worker_first_e.inc" -->
	</td></tr>
	</FORM>	
	<td align=center>
		<table style="">
		<tr align=center>
		<td width=150 style="background-color:#FF79FF;"><A Href="" target="_blank" title="�ΦL�ӽ�" style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">�ΦL</a></td>
		<td width=150 style="background-color:#FFCCCC;"><A Href='./00_01_build.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">�س]��</A></td>
		<td width=150 style="background-color:#FFDDAA;"><A Href='./00_02_sales.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">�~�ȳ�</A></td>
		<td width=150 style="background-color:#FFFFBB;"><A Href='./00_03_manager.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">�޲z��</A></td>
		<td width=150 style="background-color:#CCFF99;"><A Href='./00_04_finance.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">�]�ȳ�</A></td>
		<td width=150 style="background-color:#BBFFEE;"><A Href='./00_05_law.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">�k�ȳ�</A></td>
		<td width=150 style="background-color:#99FFFF;"><A Href='./00_06_mis.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">��T��</A></td>
		<td width=150 style="background-color:#CCCCFF;"><A Href='./00_07_golf.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">������</A></td>
		<td width=150 style="background-color:#FFB3FF;"><A Href='./00_08_fundation.asp' target=_blank style="color:#000000;text-decoration:none;font-family:'�L�n������';font-weight:bold;">�� ��</A></td>
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
	     document.write("<TR ",thcol,"><TH COLSPAN=7><font style='font-size:12pt;'>","����"+ROCYear+"�~",names[Month],thisDay,"��","</font></th>");
	     document.write("<TR ",trcol,"><TH><font style='font-size:11pt;'>��</font></TH><TH><font style='font-size:11pt;'>�@</font></TH><TH><font style='font-size:11pt;'>�G</font></TH><TH><font style='font-size:11pt;'style='font-size:11pt;'>�T</font></TH><TH><font style='font-size:11pt;'>�|</font></TH><TH><font style='font-size:11pt;'>��</font></TH><TH><font style='font-size:11pt;'>��</font></TH></TR>");
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
	//if (thisYear < 2000) thisYear=thisYear+13;
	Calendar(thisMonth,thisYear);
	
	//-->
	</SCRIPT>
	<img src="./img/demo.png" style="height:25px;vertical-align:middle;" onclick="demo()" title="�n�J�i�ܪ�����">

</TD>
</TR></TABLE>

<!-- #Include file = "./�s������.asp" -->



</center>
<script language=vbscript>
sub demo() '�i�ܪO
	MyVar = MsgBox ("�T�w�i�J�i�ܪ��������I�I�C", 64+1, "MsgBox Example")
		   if MyVar =1 then
		   	'�T�w�i�J�i�ܪ�������
		   	location.href=""
		   end if	
end sub
</script>
</BODY>
</HTML>






