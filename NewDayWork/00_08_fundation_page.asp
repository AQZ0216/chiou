<%@ Language=VBScript CODEPAGE=950 %>
<%
p_uid=request("p_uid")
if p_uid="" then
	str_url="./00_08_fundation.asp"
  response.redirect(str_url)
end if
%>
<HTML>
<HEAD>
<title>����|�t�έ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="icon" href="../daywork/img/khouse.ico" type="image/ico" />
<link rel="stylesheet" type="text/css" 
href="base_first.css" title="style1">
<style type="text/css"><!--
.tde{
	font-family:'�з���';
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
<BODY topmargin=15 >
<CENTER>
<TABLE BORDER=1 cellspacing=0 cellpadding=0 width=783>
			<tr style="height:35pt;">
				<td align=center style="background-color:#FFB3FF;"><font style="font-size:20pt;letter-spacing:2mm;">�w��i�J����|�޲z�t��!!</font></td>
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
						<td class=tde ><A class=tda Href='../ctelmsg/ct_login.asp' target='_blank' >�q�ܬ���</A></td>
						<td class=tde ><A class=tda Href='http://192.168.0.10/chiou/cbooks/bs_main.asp' target='_blank' >�خѺ޲z</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
					</tr>
					<tr align=center style="height:25pt;" >
						<td class=tde ><A class=tda Href='../transcript/ts_main.asp' target='_blank' >�����å�</A></td>
						<td class=tde ><A class=tda Href='../kcrbase/kcrbase_main.asp' target='_blank' >����Ȥ�</A></td>
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
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
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
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='../customer/login.asp' target=_top>�Ȥ���</A></td>
						<td class=tde ><A class=tda Href='../dataman/�򥻽d��.asp' target='_blank' >�򥻽d����</A></td>
						<td class=tde ><A class=tda Href='../daywork/firstpage.asp' target='_blank' >�^����</A></td>
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
	</TD>
</TR>
</TABLE>

</center>

</BODY>
</HTML>






