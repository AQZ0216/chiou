<% @codepage=950%>
<!-- �}�Ҥu�@�H���}�C -->
<!-- #Include file = "./include/array_worker_crew.inc" -->
<%
	'Ū���H���m�W
	worker = Session("worker")
	wk_order=worker
%>

<html>
<head>
<title>�d�߸��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{font-family:'�з���';background-color:'#fafad2'}
td{font-size:4.5mm;}
input.imenu { 
	font-size:4mm;				/*�r��j�p*/
	cursor:hand;				/*��ЧΦ�*/ 
	width:100%;
	height:100%;
	background-color:'#ffdab9'; 		/*�~���C��*/ 
	margin:0 0 0 0;		/*��t�W�U���k*/
     }
.sel1 { 
	font-size:4mm;				/*�r��j�p*/
	cursor:hand;				/*��ЧΦ�*/ 
	width:100%;
	height:100%;
	background-color:'#ffffee'; 		/*�~���C��*/ 
	margin:0 0 0 0;		/*��t�W�U���k*/
     }

TD.SOME{
font-family: �s�ө���;
font-size: 3.5mm;
line-height: 18px;
color:blue;
font-weight:bold;
}
TD.myd{
font-family: �s�ө���;
font-size: 3.5mm;
line-height: 18px;
}

--></style>
</head>
<body >
<center>
<form name="form1" method=post action="zwk2google_qlist.asp">
<font size=4 color="blue">�d�ߡi�u�@���ءj�ରGoogle���csv�ɮ�</font><br>
	<input type="button" name="sentb" class="cbutton" value="�T�w�d��" onclick="Verify_chk()" >
	<input type="reset" name="reset" class="cbutton" value="�M�����"  >
	<input type=button name=giveup class="cbutton" value="�^�W�@��" onclick="history.back()" >
<hr>
<table border=1 cellspacing=0 cellpadding=0 style="width:600px;">
<col style="width:100px;color:#ff0000;">
<col style="width:100px;">
<col style="width:100px;color:#ff0000;">
<col style="width:300px;">
	<tr> 
	<td align="right">���|�H���G</td>
	<td align="left">
		<select name="p_wk_doer" >
		<option value="����">�����H��</option>
<%	
	for i=1 to worker_no
		if worker_a(i-1)=worker then
			str_sel="selected"
		else
			str_sel=""
		end if
%>
		<option value="<%=worker_a(i-1)%>" <%=str_sel%>><%=worker_a(i-1)%></option>
<%

	next
%>
		</select>
	</td>	
	<td align="right">�u�@�D���G</td>
	<td align="left" >
	<input type="text" name="p_wk_item" style="width:100%;" >
	</td>
</tr>
<tr>
	<td align="center">������</td>	
	<td align="left" colspan=5>
		<input type="text" name="p_doing_date1a" style="width:100px;" >
		<img align='top' onmousedown="Cal('p_doing_date1a')" src='img/cal3.gif' width="16" height="16" align='top' style='cursor:hand'>
		�ذ�������
		<input type="text" name="p_doing_date1b" style="width:100px;" >
		<img align='top' onmousedown="Cal('p_doing_date1b')" src='img/cal3.gif' width="16" height="16" align='top' style='cursor:hand'>
	</td>	
</tr>
	<tr> 
	<td align="right">����H���G</td>
	<td align="left">
		<select name="p_wk_exe" >
		<option value="����">�����H��</option>
<%	
	for i=1 to worker_no
		'if worker_a(i-1)=worker then
		'	str_sel="selected"
		'else
			str_sel=""
		'end if
%>
		<option value="<%=worker_a(i-1)%>" <%=str_sel%>><%=worker_a(i-1)%></option>
<%

	next
%>
		</select>
	</td>
		<td align="right" colspan=2>&nbsp;</td>
</tr>
</table>
<hr>

</form>

<!--��䲣�ͪ���m-->
<Span ID=ShowCal style="position:absolute;z-index:1"></Span>

<Script Language=VBScript>
<!--
Sub Cal(TObject)

'���ͤ��A�H���Ѫ���������
 Call GetCal(Year(Now()),Month(Now()),TObject)

'�վ�<Span>����m
 ShowCal.style.left=window.event.clientX
 ShowCal.style.top=window.event.clientY
End Sub 

Sub GetCal(SYear,SMonth,TObject)
'�����Y����
Str=Str &"<table ALIGN='CENTER' BORDER='1' CELLSPACING='0' CELLPADDING='2' BGCOLOR='White' BORDERCOLOR='Gray'>"
Str=Str &"<tr><td>"
Str=Str &"        <table WIDTH='140' BORDER='0' CELLPADDING='1' CELLSPACING='0' BGCOLOR='#FFFFFF'>"
Str=Str &"                <tr HEIGHT='18' BGCOLOR='Silver'>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/prev3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='�W�@��' style='cursor:hand' OnClick='PreMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/Next3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='�U�@��' style='cursor:hand' OnClick='NextMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='100' COLSPAN='4' ALIGN='CENTER' VALIGN='MIDDLE' CLASS='SOME'><Span ID=SelYear>" & SYear & "</span>�~<Span ID=SelMon>" & SMonth & "</Span>��</td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='RIGHT' VALIGN='MIDDLE'><img SRC='img/cdia3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='��������' style='cursor:hand' OnClick='Closedate()'></td>"
Str=Str &"                </tr>"
Str=Str &"          <tr HEIGHT='15' BGCOLOR='Aliceblue'>"
Str=Str &"                <td Colspan=7>"
Str=Str &"                 <Table Border=0>"
Str=Str &"                  <tr>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>��</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>�@</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>�G</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>�T</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>�|</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>��</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>��</td>"
Str=Str &"         </tr>"
Str=Str &"        </Table>"
Str=Str &"          </tr>"
Str=Str &"          <tr>"
Str=Str &"           <td HEIGHT='1' ALIGN='MIDDLE' COLSPAN='7'><img SRC='images/line.gif' HEIGHT='1' WIDTH='140' BORDER='0'></td>"
Str=Str &"          </tr>"
Str=Str &"      <tr>"

'�Ӥ�Ĥ@�Ѥ��P��
SDate=DateValue(SYear & "/" & SMonth & "/1")
SWeek=WeekDay(SDate)

'�Ӥ�̫�@�Ѫ����
EDay=Day(DateSerial(SYear,SMonth+1,0))
EDate=DateValue(SYear & "/" & SMonth & "/" & EDay)

'�Ӥ�̫�@�Ѥ��P��
EWeek=WeekDay(EDate)

'�W��̫�@�Ѫ����
PreEDay=Day(DateSerial(SYear,SMonth,0))

'���ͤW�Ӥ몺����
Dim i
For i=1 to SWeek-1
 Str=Str & "<td CLASS='myd' width=20 align=right><font color=gray>" & PreEDay-SWeek+i+1 & "</font></td>"
Next

If SWeek=1 then
 Str=Str & "</tr>"
End if

'���ͥ��볡��
SW=SWeek
i=1
For i=1 to EDay
 '�վ㤻��r�����C��
 Select Case SW
 Case 1
  FColor="Red"
 Case 7
  FColor="Green"
 Case Else
  FColor="Black"
 End Select
  
  GDate=SYear & "/" & SMonth & "/" & i
  Str=Str & "<td CLASS='myd' width=20 align=right style='cursor:hand' onMouseOver=""this.style.backgroundColor='#FF99FF'"" onMouseOut=""this.style.backgroundColor='White'"" Onclick=""SendDate('" & GDate & "','" & TObject & "')""><Font Color=" & FColor & ">" & i & "</Font></td>"
 
'���ͤU�Ӥ볡��
SW=SW+1
 IF SW>7 then
  Str=Str & "</tr><tr>"
  SW=1
 End if
Next

J=1
For i=SW to 7
 Str=Str & "<td CLASS='myd' width=20 align=right><Font Color=Gray>" & j & "</Font></td>"
 J=j+1
Next


Str=Str & "      </tr>"
Str=Str & "</Table>"
'�N��ƤޤJ<Span>
ShowCal.InnerHTML=Str

End Sub

'�e���@�Ӥ�
Sub PreMon(TObject)
 SYear=Int(SelYear.OuterTEXT)
 SMon=int(SelMon.outerTEXT)-1

 '�P�_�O�_���e�դ@�~
 IF SMon<1 then
  SMon=12
  SYear=SYear-1
 End if
 Call GetCal(SYear,SMon,TObject)
End Sub

'�Ჾ�@�Ӥ�
Sub NextMon(TObject)
 SYear=Int(SelYear.OuterTEXT)
 SMon=int(SelMon.outerTEXT)+1

 '�P�_�O�_���e���@�~
 IF SMon>12 then
  SMon=1
  SYear=SYear+1
 End if
 Call GetCal(SYear,SMon,TObject)
End Sub

'�N��ưe�J��줺
Sub SendDate(GDate,TObject)
 document.all.namedItem(TObject).Value=GDate
 ShowCal.InnerHTML=""
End Sub

'��������e��
Sub CloseDate()
 ShowCal.InnerHTML=""
End Sub
-->
</Script>
<script language=vbscript>
sub Verify_chk()
		chkq=msgbox("�T�w�d�ߡI�I",64+1,"�T�{�T��")
		if chkq=1 then
	    document.form1.submit
	  end if
end sub
Sub add_chk()	'�s�W�k�H
   MyVar = MsgBox ("�T�w�k�H�s�W�I�I"& chr(13) & pp_data , 64+1, "MsgBox Example")
   if MyVar =1 then
   	'�T�w�s�W
   	str_nexturl="./a01_cop_add.asp"
		location.href=str_nexturl
   else
   end if
End Sub
</script>
</center>
</body>
</html>
