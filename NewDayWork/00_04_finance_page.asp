<%@ Language=VBScript CODEPAGE=950 %>
<%
p_uid=request("p_uid")
if p_uid="" then
	str_url="./00_04_finance.asp"
  response.redirect(str_url)
end if

%>
<%
' �s��Access��Ʈw../daywork/database/daywork.mdb
DBpath_acr=Server.MapPath("./database/crew.mdb")
strCon_acr="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_acr
'�إ߸�Ʈw�s������
set conDB_acr= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB_acr.Open strCon_acr
'�}�Ҹ�ƪ�W��
tb_name_acr="crew"
'�إ߸�Ʈw�s������	
set rstObj_acr=Server.CreateObject("ADODB.Recordset")
strSQL_acr="Select * from " & tb_name_acr &" where wkr_id = " & p_uid
rstObj_acr.open strSQL_acr,conDB_acr,1,1
'�p�����`��	
staff_no=rstObj_acr.recordcount
rstObj_acr.MoveFirst
for icr=1 to staff_no
	staff_name=rstObj_acr.fields("wkr_name") '�ʺ�
'����U�@���O��		
	rstObj_acr.MoveNext		
next
'������ƶ�
rstObj_acr.Close
'���]����ܼ� 
set rstObj_acr=Nothing
'������Ʈw 
conDB_acr.Close
'���]�����ܼ� 
set conDB_acr=Nothing
%>

<HTML>
<HEAD>
<title>�]�ȳ��t�έ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
<BODY topmargin=15>
<CENTER>
<TABLE BORDER=1 cellspacing=0 cellpadding=0 width=783>
			<tr style="height:35pt;">
				<td align=center style="background-color:#CCFF99;"><font style="font-size:20pt;letter-spacing:2mm;">�w��i�J�]�ȳ��޲z�t��!!</font></td>
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
						<td class=tde ><A class=tda Href='' target=_blank>�~�Z���</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >�P��ץ�</A></td>
						<td class=tde ><A class=tda Href='' target=_blank>�N�Ѯץ�</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >�N�ѷ~�ȰO��ï</A></td>
					</tr>
					<tr align=center style="height:25pt;" >

						<td class=tde ><A class=tda Href='' target='_blank' style="color:red;background-color:#e1e1e1;" >�i�~�Z����j</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' style="color:red;background-color:#e1e1e1;" > �i�ץ�]�֡j</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
					</tr>
					<tr align=center style="height:25pt;" >
						<td class=tde ><A class=tda Href='../expenses/a01_main.asp' target='_blank' style="color:red;background-color:#e1e1e1;" > �i��X�ҩ���j</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >&nbsp;</A></td>
					</tr>
					<tr align=center style="height:25pt;" >
						<td class=tde ><A class=tda Href='../rights/a00_login.asp' target='_blank' style="color:red;background-color:#e1e1e1;" > �i�v���޲z�j</A></td>
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
						<td class=tde ><A class=tda Href='' target='_blank' title="��j���ں޲z�t��">��j���ں޲z</A></td>
						<td class=tde ><A class=tda Href='' target=_top>�Ȥ���</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >�򥻽d����</A></td>
						<td class=tde ><A class=tda Href='' target='_blank' >�^����</A></td>
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






