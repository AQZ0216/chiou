<%
' 連結Access資料庫./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="linkdata"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " order by lk_row asc, lk_col asc"
rstObj1.open strSQL_show,conDB,3,1
'計算資料總數	
totalput=rstObj1.recordcount
if totalput=0 then
else
%>
<table border="1" cellspacing=0 cellpadding=0 width=783>
<col span=6 style="width:16.6%;">
<%
      '移至第一筆資料
      rstobj1.MoveFirst
      p_04old=0
      '列出資料項目
      for i=1 to totalput
      	'設定空白資料之反映
      p_id=rstObj1.fields("lk_id")		'id	
      p_01=rstObj1.fields("lk_url")		'連結網址
      p_02=rstObj1.fields("lk_item")		'短標題
      p_03=rstObj1.fields("lk_title")		'描述
      p_04=rstObj1.fields("lk_row")		'列
      p_05=rstObj1.fields("lk_col")		'欄


if p_04=p_04old then
else
      if p_04<>1 then response.Write   "</tr>"
      response.Write   "<tr align=center style='height:20pt;' >"
      p_04old=p_04
end if

if p_02="" or isnull(p_02) then
   response.write "<td>&nbsp;</td>"
else
   if len(p_02)>7 then
      str_ft="font-size:11pt;"
   else
      str_ft="font-size:12.5pt;"
   end if
   if date()>="2018/11/9" and date()<="2018/12/9" then
	 	if p_02="匠人須知投票" then str_ft="font-size:12.5pt;background-color:#ffaeff;"
'	 	if p_02="分機表" then str_ft="font-size:12.5pt;font-weight:bold;background-color:#ffaeff;"
	 end if
%>
<td class=urlcmd title='<%=p_03%>' style='<%=str_ft%>'><A Href='<%=p_01%>' target='_blank' ><%=p_02%></A></td>
<%

end if

      '移到下一筆記錄
      rstObj1.MoveNext
      if rstObj1.EOF=True then exit for
      next

   response.Write   "</tr>"
end if

'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing
%>
	
</table>
