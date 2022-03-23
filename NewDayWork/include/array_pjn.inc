<%
'工作人員陣列daywork.mdb worker_data
dim pjn_a()
dim pjnid_a()
%>
<%
' 連結Access資料庫daywork.mdb
DBpath_a1=Server.MapPath("./database/daywork.mdb")
strCon_a1="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_a1
'建立資料庫連結物件
set conDB_a1= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB_a1.Open strCon_a1
'開啟資料表名稱
tb_name_a1="project_data"
'建立資料庫存取物件	
set rstObj_a1=Server.CreateObject("ADODB.Recordset")
strSQL_a1="Select * from " & tb_name_a1
rstObj_a1.open strSQL_a1,conDB_a1,1,1
'計算資料總數	
pjn_no=rstObj_a1.recordcount
if pjn_no=0 then
else
      '重設陣列數目
      redim pjn_a(Cint(pjn_no))
      redim pjnid_a(Cint(pjn_no))
      rstObj_a1.MoveFirst
      for i=1 to pjn_no
      	pjn_a(i-1)=rstObj_a1.fields("pj_02")
      	pjnid_a(i-1)=rstObj_a1.fields("pj_id")
      '移到下一筆記錄		
      	rstObj_a1.MoveNext		
      next
end if
'關閉資料集
rstObj_a1.Close
'重設資料變數 
set rstObj_a1=Nothing
'關閉資料庫 
conDB_a1.Close
'重設物件變數
set conDB_a1=Nothing 
%>


	