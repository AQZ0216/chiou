<%
'群組陣列 staffgp_a()
dim staffgp_a()
' 連結Access資料庫../holiday/database/crew.mdb
DBpath_acr=Server.MapPath("../holiday/database/crew.mdb")
strCon_acr="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_acr
'建立資料庫連結物件
set conDB_acr= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB_acr.Open strCon_acr
'開啟資料表名稱
tb_name_acr="wk_sgp"
'建立資料庫存取物件	
set rstObj_acr=Server.CreateObject("ADODB.Recordset")
strSQL_acr="Select * from " & tb_name_acr &" order by gp_item asc"
rstObj_acr.open strSQL_acr,conDB_acr,1,3
'計算資料總數	
staff_no=rstObj_acr.recordcount
'重設陣列數目
redim staffgp_a(Cint(staff_no))
rstObj_acr.MoveFirst
for icr=1 to staff_no
	staffgp_a(icr-1)=rstObj_acr.fields("gp_item") '群組
'移到下一筆記錄		
	rstObj_acr.MoveNext		
next
'關閉資料集
rstObj_acr.Close
'重設資料變數 
set rstObj_acr=Nothing
'關閉資料庫 
conDB_acr.Close
'重設物件變數 
set conDB_acr=Nothing
%>
