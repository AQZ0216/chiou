<%@ Language=VBScript CODEPAGE=950 %>

<%
   '弄﹎
   worker = Session("worker")

%>
<%
'ヰ安戈
function hd_man(p_hdate)
   pstr_hdman =""
    ' 硈挡Access戈畐holiday.mdb
    DBpath_fh=Server.MapPath("../holiday/database/holiday.mdb")
    strCon_fh="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fh
    'ミ戈畐硈挡ン
    set conDB_fh= Server.CreateObject("ADODB.Connection")
    '硈挡戈畐	
    conDB_fh.Open strCon_fh
    '秨币戈嘿
    tb_name_fh="ヰ安灿"
	'ミ戈畐ン
	set rstObj1_fh=Server.CreateObject("ADODB.Recordset")
	strSQL_show_fh="Select * from " & tb_name_fh & " where ヰ安ら = #"& p_hdate &"# order by 安id asc "
	rstObj1_fh.open strSQL_show_fh,conDB_fh,3,1
	totalput_fh=rstObj1_fh.recordcount
if not rstObj1_fh.EOF then
	rstObj1_fh.Movefirst
	for i = 1 to totalput_fh
		hd_id=rstObj1_fh.fields("hd_id")
		icon_id=rstObj1_fh.fields("安id")
		hd_hrs=rstObj1_fh.fields("ヰ安计")
		hd_check=rstObj1_fh.fields("絋粄")
		hd_man=rstObj1_fh.fields("﹎")'﹎
		hd_img=left(rstObj1_fh.fields("安嘿"),1)
		hd_cname=right(rstObj1_fh.fields("安嘿"),len(rstObj1_fh.fields("安嘿"))-1)
		'∕﹚安肅︹
		select case icon_id
'		   Case 1  f_color = "#FF0000"    '〕そヰ
'		   Case 2  f_color = "#000000"    '《ㄆ安
'		   Case 3  f_color = "#0000FF"    '◎痜安
'		   Case 4  f_color = "#BB5500"    '±そ安
'		   Case 5  f_color = "#000000"    '』赤安
'		   Case 6  f_color = "#00FF00"    '〉安
'		   Case 7  f_color = "#FF0088"    '〗疭ヰ
'		   Case 8  f_color = "#EE7700"    '」玻安
'		   Case 9  f_color = "#BBBB00"    '》盉安
'		   Case 15  f_color = "#000000"   '“ゼゴ
'		   Case 16  f_color = "#FF0000"   '【逼痁
'		   Case Else   f_color = "#FF0000"
		   Case 1  f_color = "#000000"    '〕そヰ
		   Case 2  f_color = "#000000"    '《ㄆ安
		   Case 3  f_color = "#000000"    '◎痜安
		   Case 4  f_color = "#000000"    '±そ安
		   Case 5  f_color = "#000000"    '』赤安
		   Case 6  f_color = "#000000"    '〉安
		   Case 7  f_color = "#000000"    '〗疭ヰ
		   Case 8  f_color = "#000000"    '」玻安
		   Case 9  f_color = "#000000"    '》盉安
		   Case 15  f_color = "#000000"   '“ゼゴ
		   Case 16  f_color = "#000000"   '【逼痁
		   Case Else   f_color = "#000000"
		End Select
		if icon_id=1 or icon_id=15 then
		    if icon_id=15 then
		       pstr_hdman = pstr_hdman & "<font style='font-size:15px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man & "</font><br>"
   	           end if
		else
		  pstr_hdman = pstr_hdman & "<font style='font-size:15px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man &"("& hd_hrs&")&nbsp;</font><br>"
		end if
		    'Response.Write "</font><br>"
		rstObj1_fh.MoveNext
		if rstObj1_fh.EOF=true then exit for
	next
else
end if
	'闽超戈栋
	rstObj1_fh.Close
	'砞戈跑计 
	set rstObj1_fh=Nothing
    '闽超戈畐
    conDB_fh.Close
    '砞ン跑计 
    set conDB_fh=Nothing
  hd_man=pstr_hdman
end function
%>
<%
'琩高琌Τン
Function exist_attach(pwk_id)
      ' 硈挡Access戈畐daywork.mdb
      DBpath_fe=Server.MapPath("./database/attach_file.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      'ミ戈畐硈挡ン
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '硈挡戈畐	
      conDB_fe.Open strCon_fe
      '秨币戈嘿
      tb_name_fe="file_data"
      'ミ戈畐ン	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where del_ok = false and wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '闽超戈栋
      rstObj1_fe.Close
      '砞戈跑计
      set rstObj1_fe=Nothing
      '闽超戈畐 
      conDB_fe.Close
      '砞ン跑计
      set conDB_fe=Nothing
      exist_attach=totalput_fe
End Function

%>
<%
'砞﹚跑计 
dim dbConn, rs, nDex, nMonth, nYear, dtDate

' Get the current date
dtDate = Now()

' Set the Month and Year
   nWeeks=request("nWeeks")             '㏄计弄ㄏノ
   nYear=request("nYear")             '弄ㄏノ
if nWeeks = "" then nWeeks =DatePart("ww",Date()) '弄秅计
if nYear = "" then nYear = Year(Date())
	nMonth=month( Now())   
   '砞﹚秅计ぇ材ぱwkn_1 の程ぱwkn_2
   date00=dateserial(nYear,1,1)  'ぇ材ぱ
  dayweek_2f=DateAdd("d",7-DatePart("w",date00)+1,date00)   '材秅ぇ材ぱ

   wkn_1=DateAdd("d",(nWeeks-2)*7 ,dayweek_2f)
   wkn_2=DateAdd("d",((nWeeks-2)*7+6) ,dayweek_2f)
   wkn_0=DateAdd("d",-7 ,wkn_1)


'nMonth = Request("nMonth")
'nYear = Request("nYear")
'if nMonth = "" then nMonth = Month(dtDate)
'if nYear = "" then nYear = Year(dtDate)

select case cint(Weekday(date()))
case 1
   cswday="琍戳ら"
case 2
   cswday="琍戳"
case 3
   cswday="琍戳"
case 4
   cswday="琍戳"
case 5
   cswday="琍戳"
case 6
   cswday="琍戳き"
case 7
   cswday="琍戳せ"

end select

'砞﹚琍戳肅︹ 
bgc1="#F0FFF0"    '睭独︹lightyellow 
bgc6="#F0FFF0" 'lightskyblue
bgc7="#F0FFF0" 'lightgreen

' Set the date to the first of the current month
'dtDate = DateSerial(nYear, nMonth, 1)
dtDate=wkn_1

if int(nMonth)<10 then
   strnMonth="0"&cstr(nMonth)
else
   strnMonth=cstr(nMonth)
end if
dcodeym=cstr(nYear)&strnMonth

'砞﹚session("strbackURL")
'strbackURL="wk_week_now.asp?nWeeks="&nWeeks&"&nYear="&nYear
'session("strbackURL")=strbackURL

%>
<%
'砞﹚る 
if nMonth = 1 then 
   pre1month=12
   pre1year=nYear-1
else
   pre1month=nMonth-1
   pre1year=nYear
end if
pre2month=nMonth
pre2year=nYear-1
pre3month=nMonth
pre3year=nYear+1
if nMonth = 12 then 
   pre4month=1
   pre4year=nYear+1
else
   pre4month=nMonth+1
   pre4year=nYear
end if
%>
<%
pstart=dateserial(nYear,nMonth,1)
pend=dateadd("m",1,pstart)
'琩高セる琌ミEAD穦某         p_wk_item="08:20-09:00 EAD穦某"
function find_ead(pstart,pend)
      ' 硈挡Access戈畐daywork.mdb
      DBpath_ead=Server.MapPath("./database/daywork.mdb")
      strCon_ead="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_ead
      'ミ戈畐硈挡ン
      set conDB_ead= Server.CreateObject("ADODB.Connection")
      '硈挡戈畐   
      conDB_ead.Open strCon_ead
      '秨币戈嘿
      tb_name_ead="work_data"
      'ミ戈畐ン	
      set rstObj1_ead=Server.CreateObject("ADODB.Recordset")
      strSQL_show_ead="Select * from " & tb_name_ead & " where wk_item like '08:20-09:00 EAD穦某' and wk_order like '紌' and doing_date1 >= #"& pstart &"# and doing_date1 < #"& pend &"# order by doing_date1 asc"
      rstObj1_ead.open strSQL_show_ead,conDB_ead,3,1
      totalput_ead=rstObj1_ead.recordcount
      '闽超戈栋
      rstObj1_ead.Close
      '砞戈跑计 
      set rstObj1_ead=Nothing
      '闽超戈畐 
      conDB_ead.Close
      '砞ン跑计 
      set conDB_ead=Nothing
      find_ead=totalput_ead
end function
pchk_ead=find_ead(pstart,pend)

%>

<!-- 秨币戈畐 -->
<%
' 硈挡Access戈畐daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'ミ戈畐硈挡ン
set conDB= Server.CreateObject("ADODB.Connection")
'硈挡戈畐   
conDB.Open strCon
'秨币戈嘿
tb_name="work_data"
%>

<HTML>
<HEAD>
<title>妓狾夹肈</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css"><!--
body{
   margin:10px 0 0 0;      /*娩絫オ*/
   font-family:'穝灿砰';      /**/
   font-size:12.5px;          /*砰*/
   background-color:'#F0FFF0'; /*璉春肅︹*/
   /*letter-spacing:2px;  */
   }
notetext{
   font-family:'穝灿砰';      /**/
   font-size:12.5px;          /*砰*/
   }
daytext{
   font-family:'穝灿砰';      /**/
   font-size:12.5px;          /*砰*/
   }

/*硈挡LINKぇは莱*/
A:link{color:black}     /*ゼ硈挡ぇ砰肅︹*/
A:visited{color:black}  /*纯竒硈挡ぇ砰肅︹*/
A:active{color:black}   /*硈挡ぇ砰肅︹*/
--></style>

</HEAD>
<body>
<center>
<form method="post" name="form1" action="">
<table border=0 bgcolor="gray" cellpadding=1 style="width:100%;border-width:0px;width:970px;">
<col style="width:56%;background-color:#F0FFF0;">
<col style="width:44%;background-color:#F0FFF0;">
<tr style="height:25px">
   <td align=center style="font-size:18px;letter-spacing:1px;" >
      <b><%=worker%>ゼЧΘら句</b>
   <td colspan=1 align=center style="font-size:15px;cursor:hand;">
   さぱ琌&nbsp;﹁じ<%=cstr(Year(date()))%><%=cstr(Month(date()))%>る<%=cstr(Day(date()))%>ら&nbsp;<%=cstr(cswday)%> 
   <br>
   陪ボ<%=nYear%>
   陪ボ秅计<%=nWeeks%>
</table>
   <table border=0 bgcolor="gray" cellpadding=1 style="width:100%;border-width:0px;width:970px;">
   <col style="width:14%;background-color:<%=bgc7%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc1%>;">
   <col style="width:14%;background-color:<%=bgc6%>;">
   <tr style="height:45px;">
      <td align=center><font color="black"><b>琍戳ら<br>Sunday</b></font></td>
      <td align=center><font color="black"><b>琍戳<br>Monday</b></font></td>
      <td align=center><font color="black"><b>琍戳<br>Tuesday</b></font></td>
      <td align=center><font color="black"><b>琍戳<br>Wednesday</b></font></td>
      <td align=center><font color="black"><b>琍戳<br>Thursday</b></font></td>
      <td align=center><font color="black"><b>琍戳き<br>Friday</b></font></td>
      <td align=center><font color="black"><b>琍戳せ<br>Saturday</b></font></td></tr>
   <tr bgcolor="#ffffc0" style="height:60px;">
      <% 
      ' Add blank cells until the proper day糤﹚竚 
      for nDex = 1 to Weekday(dtDate) - 1
         Response.Write "<td bgcolor=""#c0c0c0"">&nbsp;</td>"
      next
      '秨﹍恶ら戳 
      do
         'ら戳锣传 
         if int(Day(dtDate))<10 then
            strnDay="0"&cstr(Day(dtDate))
         else
            strnDay=cstr(Day(dtDate))
         end if
         dcodeymd=dcodeym&strnDay
         dcodeymd_a=dtDate
         '砞﹚肅︹ 
         select case cint(Weekday(dtDate))
         case 1
            bgc=bgc7
         case 7
            bgc=bgc6
         case else 
            bgc=bgc1 
         end select

         'ら戳ず恶糶ゅ 
         if int(nYear)=int(Year(date()))and int(nMonth)=int(Month(date()))and int(Day(dtDate))=int(day(date())) then
            Response.Write "<td valign=""top"" bgcolor=""#f0ffae"" >"
         else        
            Response.Write "<td valign=""top"" bgcolor="&bgc&" >"
         end if
               Response.Write "<table border=0 cellpadding=0 style=""width:100%;height:100%;border-width:0px;"">"
               Response.Write "<tr style=""width:100%;height:10px;"">"
               Response.Write "<td align=center style=""width:50%;background-color:#dcdcdc;font-size:13px;"">"
               Response.Write Day(dtDate)
               Response.Write "<td align=center style=""width:50%;background-color:#dcdcdc;"">"
               Response.Write "<font style=""font-size:13px;""><a href='wk_add.asp?datecode="&dtDate&"'>穝糤"
               Response.Write "</a></font>"
               'Response.Write "<font style=""text-size:3mm;""><b>&nbsp;"
               'Response.Write "</b></font>"
               Response.Write "<tr><td colspan=2 align=left valign=top>"
         '恶ㄆン戈 
            'Response.Write "<font size=""-1""><b>" &dcodeymd&"</b></font><br>"
            'Response.Write "<font size=""-1""><b>" &totalput&"</b></font><br>"
            'ミ戈畐ン  
            set rstObj1=Server.CreateObject("ADODB.Recordset")
            strSQL_show="Select * from " & tb_name & " where doing_date1 = #"&dcodeymd_a&"# and wk_undoer like '%"&worker&"%' order by wk_item asc , wk_id asc"
            rstObj1.open strSQL_show,conDB,3,1
            totalput=rstObj1.recordcount
         if not rstObj1.EOF then
            rstObj1.Movefirst
            for i = 1 to totalput
            	 wk_headline=rstObj1.fields("headline")  '禲皑縊
               wk_id=rstObj1.fields("wk_id")
               '浪琩琌Τン exist_attach(wk_id)
               attach_no=exist_attach(wk_id)
               if attach_no=0 then
                  str_colors="color:#000000;"
               else
                  str_colors="color:#0000FF;"
               end if
               if rstObj1.fields("wk_password")="" or isnull(rstObj1.fields("wk_password")) then
               else
                  str_colors="color:#0000FF;"
               end if
								p_nexe=rstObj1.fields("wk_exe")	'磅︽
								if Instr(1, p_nexe, worker, 1)>0 or Instr(1, p_nexe, "砰", 1)>0 then
									str_bgc="background-color:#99FF99;"	
								else
									str_bgc=""
								end if
               Response.Write "<font style='font-size:13px;"& str_bgc & str_colors &"' >" & i & "<a href='wk_show.asp?wk_id="&wk_id&"' style='letter-spacing:1.5pt;font-size:11pt;"& str_colors &"' >" & replace (rstObj1.fields("wk_item"),"キ底冻","<font color=fuchsia >キ底冻</font>")
               Response.Write "</a></font>"
if wk_headline>5 then response.write "<img src='./img/gnome_chess.png' title='禲皑縊癟' width=19 style='vertical-align:top;'>"                
               Response.Write "<br>"
               rstObj1.MoveNext
               if rstObj1.EOF=true then exit for
            next
         else
         end if
            '闽超戈栋
            rstObj1.Close
            '砞戈跑计 
            set rstObj1=Nothing
             str_hdman=hd_man(dcodeymd_a)     'ヰ安戈
            response.write str_hdman
            Response.Write "</table>"
            Response.Write "</td>"
                  
            if WeekDay(dtDate) = 7 then 
               Response.Write "</tr>" & vbCrLf & "<tr  bgcolor=""#ffffc0"" style=""height:60px;"">"
            end if
            dtDate = DateAdd("d", 1, dtDate)
            
      loop until ( dtDate > wkn_2 )
'      loop until (Month(dtDate) <> CInt(nMonth))
      'ら戳恶糶ЧΘ 
         ' Add blank cells to fill out the rest of the month if needed糤﹚竚
         if Weekday(dtDate) <> 1 then 
            for nDex = Weekday(dtDate) to 7
               Response.Write "<td bgcolor=""#C0C0C0"">&nbsp;</td>"
            next
         end if
         %>
      </tr>
      <tr>
      <td colspan=7>
<b>
<font color="#FF0000">〕そヰ</font>
<font color="#000000">《ㄆ安</font>
<font color="#0000FF">◎痜安</font>
<font color="#BB5500">±そ安</font>
<font color="#000000">』赤安</font>
<font color="#00FF00">〉安</font>
<font color="#FF0088">〗疭ヰ</font>
<font color="#EE7700">」玻安</font>
<font color="#BBBB00">》盉安</font><BR>
<font color="#000000">“ゼゴ</font>
<font color="#FF0000">【逼痁</font>
<font color="#000000">()叫安计</font>
</b>
</td>
      </tr>
   </table>

</form>
<%

'闽超戈畐 
conDB.Close
'砞ン跑计 
set conDB=Nothing 
%>    
</center>
</body>
</html>
