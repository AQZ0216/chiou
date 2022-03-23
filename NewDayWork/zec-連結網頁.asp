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
      '移至第一筆資料
      rstobj1.MoveFirst
      '列出資料項目
      for i=1 to totalput
      	'設定空白資料之反映
      p_id=rstObj1.fields("lk_id")		'id	
      p_01=rstObj1.fields("lk_url")		'連結網址
      p_02=rstObj1.fields("lk_item")		'短標題
      p_03=rstObj1.fields("lk_title")		'描述
      p_04=rstObj1.fields("lk_row")		'列
      p_05=rstObj1.fields("lk_col")		'欄
%>
<button class="w3-button w3-large w3-pale-yellow  w3-border w3-border-brown w3-round-large " style="margin:2px;padding:3px;width:150px;" onclick="url_new('<%=p_01%>')" >
<%=p_02%>
</button>
<%
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
%>
	

