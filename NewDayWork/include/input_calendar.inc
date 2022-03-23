<!--月曆產生的位置-->
<Span ID=ShowCal style="position:absolute;z-index:1"></Span>

<Script Language=VBScript>

sub Cal(TObject)

'產生月曆，以今天的日期為基準
nYear=Year(Now())
nMonth=Month(Now())
 GetCal nYear,nMonth,TObject

'調整<Span>的位置
 ShowCal.style.left=window.event.clientX
 ShowCal.style.top=window.event.clientY
End sub 

Sub GetCal(SYear,SMonth,TObject)
'月曆抬頭部分
Str=Str &"<table ALIGN='CENTER' BORDER='1' CELLSPACING='0' CELLPADDING='2' BGCOLOR='White' BORDERCOLOR='Gray'>"
Str=Str &"<tr><td>"
Str=Str &"        <table WIDTH='140' BORDER='0' CELLPADDING='1' CELLSPACING='0' BGCOLOR='#FFFFFF'>"
Str=Str &"                <tr HEIGHT='18' BGCOLOR='Silver'>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/prev3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='上一月' style='cursor:hand' OnClick='PreMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='LEFT' VALIGN='MIDDLE'><img SRC='img/Next3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='下一月' style='cursor:hand' OnClick='NextMon(""" & TObject & """)'></td>"
Str=Str &"                        <td WIDTH='100' COLSPAN='4' ALIGN='CENTER' VALIGN='MIDDLE' CLASS='SOME'><Span ID=SelYear>" & SYear & "</span>年<Span ID=SelMon>" & SMonth & "</Span>月</td>"
Str=Str &"                        <td WIDTH='20' HEIGHT='18' ALIGN='RIGHT' VALIGN='MIDDLE'><img SRC='img/cdia3.gif' WIDTH='18' HEIGHT='18' BORDER='0' ALT='關閉視窗' style='cursor:hand' OnClick='Closedate()'></td>"
Str=Str &"                </tr>"
Str=Str &"          <tr HEIGHT='15' BGCOLOR='Aliceblue'>"
Str=Str &"                <td Colspan=7>"
Str=Str &"                 <Table Border=0>"
Str=Str &"                  <tr>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>日</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>一</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>二</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>三</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>四</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>五</td>"
Str=Str &"                    <td ALIGN='RIGHT' CLASS='SOME' WIDTH='20' HEIGHT='15' VALIGN='BOTTOM'>六</td>"
Str=Str &"         </tr>"
Str=Str &"        </Table>"
Str=Str &"          </tr>"
Str=Str &"          <tr>"
Str=Str &"           <td HEIGHT='1' ALIGN='MIDDLE' COLSPAN='7'><img SRC='images/line.gif' HEIGHT='1' WIDTH='140' BORDER='0'></td>"
Str=Str &"          </tr>"
Str=Str &"      <tr>"

'該月第一天之星期
SDate=DateValue(SYear & "/" & SMonth & "/1")
SWeek=WeekDay(SDate)

'該月最後一天的日期
EDay=Day(DateSerial(SYear,SMonth+1,0))
EDate=DateValue(SYear & "/" & SMonth & "/" & EDay)

'該月最後一天之星期
EWeek=WeekDay(EDate)

'上月最後一天的日期
PreEDay=Day(DateSerial(SYear,SMonth,0))

'產生上個月的部分
Dim i
For i=1 to SWeek-1
 Str=Str & "<td CLASS='myd' width=20 align=right><font color=gray>" & PreEDay-SWeek+i+1 & "</font></td>"
Next

If SWeek=1 then
 Str=Str & "</tr>"
End if

'產生本月部分
SW=SWeek
i=1
For i=1 to EDay
 '調整六日字型的顏色
 Select Case SW
 Case 1
  FColor="Red"
 Case 7
  FColor="Green"
 Case Else
  FColor="Black"
 End Select
  
  GDate=SYear & "/" & SMonth & "/" & i
'  Str=Str & "<td CLASS='myd' width=20 align=right style='cursor:hand' onMouseOver=""this.style.backgroundColor='#FF99FF'"" onMouseOut=""this.style.backgroundColor='White'"" Onclick=""SendDate('" & GDate & "','" & TObject & "')""><Font Color=" & FColor & ">" & i & "</Font></td>"
  Str=Str & "<td CLASS='myd' width=20 align=right style='cursor:hand' onMouseOver=""this.style.backgroundColor='#FF99FF'"" onMouseOut=""this.style.backgroundColor='White'"" Onclick=""SendDate('" & GDate & "','" & TObject & "')""><Font Color=" & FColor & ">" & i & "</Font></td>"
 
'產生下個月部分
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

'將資料引入<Span>
ShowCal.InnerHTML=Str

End Sub

'前移一個月
Sub PreMon(TObject)
 SYear=Int(SelYear.OuterTEXT)
 SMon=int(SelMon.outerTEXT)-1

 '判斷是否往前調一年
 IF SMon<1 then
  SMon=12
  SYear=SYear-1
 End if
 Call GetCal(SYear,SMon,TObject)
End Sub

'後移一個月
Sub NextMon(TObject)
 SYear=Int(SelYear.OuterTEXT)
 SMon=int(SelMon.outerTEXT)+1

 '判斷是否往前往一年
 IF SMon>12 then
  SMon=1
  SYear=SYear+1
 End if
 GetCal SYear,SMon,TObject 
End Sub

'將資料送入欄位內
function SendDate(GDate,TObject)
 document.all.namedItem(TObject).Value=GDate
 ShowCal.InnerHTML=""
End function

'關閉日期畫面
Sub CloseDate()
 ShowCal.InnerHTML=""
End Sub
</Script>
