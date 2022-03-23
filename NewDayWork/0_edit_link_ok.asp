<%@ Language=VBScript CODEPAGE=950 %>
<%

p_id = request("p_id")   'id
'設定讀取資料編號
p_01=Request("p_01")    '連結網址
p_02=Request("p_02")    '簡短標題
p_03=Request("p_03")    '描述

'讀取分區類別陣列
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
   strSQL_show="Select * from " & tb_name & " where lk_id="& p_id &" order by lk_id asc"
rstObj1.open strSQL_show,conDB,3,3
'計算資料總數	
totalput01=rstObj1.recordcount
'列出資料項目

      rstObj1.fields("lk_url")	=p_01	'連結網址
      rstObj1.fields("lk_item")=p_02		'短標題
      rstObj1.fields("lk_title")=p_03		'描述

rstObj1.UpdateBatch
'關閉資料集
rstObj1.Close
'重設資料變數
set rstObj1=Nothing
'關閉資料庫
conDB.Close
'重設物件變數
set conDB=Nothing

   str_url="firstpage_elink.asp"   '讀取列表網址
   response.redirect(str_url)


%>