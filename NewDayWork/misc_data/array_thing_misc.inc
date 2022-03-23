<%
'事件陣列misc_data.mdb thing_table
dim thing_a()
%>
<%
' 連結Access資料庫misc_data.mdb
DBpath_a=Server.MapPath("./misc_data.mdb")
strCon_a="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_a
'建立資料庫連結物件
set conDB_a= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB_a.Open strCon_a
'開啟資料表名稱
tb_name_a="thing_table"
%>
<%
'建立資料庫存取物件	
set rstObj_a=Server.CreateObject("ADODB.Recordset")
strSQL_a="Select * from " & tb_name_a 
rstObj_a.open strSQL_a,conDB_a,3
'計算資料總數	
thing_no=rstObj_a.recordcount
'重設陣列數目
redim thing_a(Cint(thing_no))
rstObj_a.MoveFirst
for i=1 to thing_no
	thing_a(i-1)=rstObj_a.fields("thing")
'移到下一筆記錄		
	rstObj_a.MoveNext		
next
'關閉資料集
rstObj_a.Close
'重設資料變數 
set rstObj_a=Nothing
'關閉資料庫 
conDB_a.Close
'重設物件變數 
set conDB_a=Nothing 
%>


	