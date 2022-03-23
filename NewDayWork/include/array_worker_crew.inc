<%
'工作人員陣列daywork.mdb worker_data
dim worker_a()
dim cworker_a()
dim eworker_a()
dim staff_a()
dim staff_gp_a()
dim staff_dp_a()
dim staff_id_a()
' 連結Access資料庫daywork.mdb
DBpath_a1=Server.MapPath("./database/crew.mdb")
strCon_a1="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_a1
'建立資料庫連結物件
set conDB_a1= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB_a1.Open strCon_a1
'開啟資料表名稱
tb_name_a1="crew"
'建立資料庫存取物件	
set rstObj_a1=Server.CreateObject("ADODB.Recordset")
'strSQL_a1="Select * from " & tb_name_a1 &" where hide = false order by wk_sgp asc, wk_gp_sq asc, wkr_id asc"
strSQL_a1="Select * from " & tb_name_a1 &" where hide = false order by wk_gp_sq asc"
rstObj_a1.open strSQL_a1,conDB_a1,3
'計算資料總數	
worker_no=rstObj_a1.recordcount
staff_no=rstObj_a1.recordcount
'重設陣列數目
redim worker_a(Cint(worker_no))
redim cworker_a(Cint(worker_no))
redim eworker_a(Cint(worker_no))
'重設陣列數目
redim staff_a(Cint(staff_no))
redim staff_gp_a(Cint(staff_no))
redim staff_dp_a(Cint(staff_no))
redim staff_id_a(Cint(staff_no))
rstObj_a1.MoveFirst
for i=1 to worker_no
	worker_a(i-1)=rstObj_a1.fields("worker")     '中文名
	cworker_a(i-1)=rstObj_a1.fields("wkr_name")     '全中文名
	eworker_a(i-1)=rstObj_a1.fields("e_name")   '英文名
	staff_id_a(i-1)=rstObj_a1.fields("wkr_id") 'id
	staff_a(i-1)=rstObj_a1.fields("e_name") '暱稱
	staff_gp_a(i-1)=rstObj_a1.fields("wk_sgp") '群組
	staff_dp_a(i-1)=rstObj_a1.fields("wk_gp") '部門	
'移到下一筆記錄		
	rstObj_a1.MoveNext		
next
'關閉資料集
rstObj_a1.Close
'重設資料變數 
set rstObj_a1=Nothing
'關閉資料庫 
conDB_a1.Close
'重設物件變數 
set conDB_a1=Nothing 

'======部門人員字串============
stra_dp01="" '總經理室
stra_dp02="" '管理部
stra_dp03="" '企劃部
stra_dp04="" '業務部
stra_dp05="" '法務+企劃部
stra_dp06="" '財務部
stra_dp07="" '資訊+管理部
stra_dp08="" '建設部
stra_dp09="" '社企基金會
stra_dp10="" '我家農業

stra_dpa1="" '業1
stra_dpa2="" '業2
stra_dpa3="" '業3
stra_dpa4="" '業4
stra_dpa5="" '業5

for ki=1 to worker_no

	if instr(1,staff_dp_a(ki-1),"總經理室",1)>0 then
			if  stra_dp01="" then
				stra_dp01= worker_a(ki-1)
			else  
				stra_dp01= stra_dp01 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"管理部",1)>0 then
			if  stra_dp02="" then
				stra_dp02= worker_a(ki-1)
			else  
				stra_dp02= stra_dp02 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"企劃部",1)>0 then
			if  stra_dp03="" then
				stra_dp03= worker_a(ki-1)
			else  
				stra_dp03= stra_dp03 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"業務部",1)>0 then
			if  stra_dp04="" then
				stra_dp04= worker_a(ki-1)
			else  
				stra_dp04= stra_dp04 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"法務部",1)>0 then
			if  stra_dp05="" then
				stra_dp05= worker_a(ki-1)
			else  
				stra_dp05= stra_dp05 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"財務部",1)>0 then
			if  stra_dp06="" then
				stra_dp06= worker_a(ki-1)
			else  
				stra_dp06= stra_dp06 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"資訊部",1)>0 then
			if  stra_dp07="" then
				stra_dp07= worker_a(ki-1)
			else  
				stra_dp07= stra_dp07 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"建設部",1)>0 then
			if  stra_dp08="" then
				stra_dp08= worker_a(ki-1)
			else  
				stra_dp08= stra_dp08 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"社企",1)>0 then
			if  stra_dp09="" then
				stra_dp09= worker_a(ki-1)
			else  
				stra_dp09= stra_dp09 & "," & worker_a(ki-1)
			end if
	end if
	if instr(1,staff_dp_a(ki-1),"我家農業",1)>0 then
			if  stra_dp10="" then
				stra_dp10= worker_a(ki-1)
			else  
				stra_dp10= stra_dp10 & "," & worker_a(ki-1)
			end if
	end if
	 
'	Select Case staff_dp_a(ki-1)
'		Case "總經理室" 
'			if  stra_dp01="" then
'				stra_dp01= worker_a(ki-1)
'			else  
'				stra_dp01= stra_dp01 & "," & worker_a(ki-1)
'			end if
'		Case "管理部"  
'			if  stra_dp02="" then
'				stra_dp02= worker_a(ki-1)
'			else  
'				stra_dp02= stra_dp02 & "," & worker_a(ki-1)
'			end if
'		Case "企劃部"  
'			if  stra_dp03="" then
'				stra_dp03= worker_a(ki-1)
'			else  
'				stra_dp03= stra_dp03 & "," & worker_a(ki-1)
'			end if
'		Case "業務部"  
'			if  stra_dp04="" then
'				stra_dp04= worker_a(ki-1)
'			else  
'				stra_dp04= stra_dp04 & "," & worker_a(ki-1)
'			end if
'		Case "法務部"  
'			if  stra_dp05="" then
'				stra_dp05= worker_a(ki-1)
'			else  
'				stra_dp05= stra_dp05 & "," & worker_a(ki-1)
'			end if
'		Case "財務部"  
'			if  stra_dp06="" then
'				stra_dp06= worker_a(ki-1)
'			else  
'				stra_dp06= stra_dp06 & "," & worker_a(ki-1)
'			end if
'		Case "資訊部"  
'			if  stra_dp07="" then
'				stra_dp07= worker_a(ki-1)
'			else  
'				stra_dp07= stra_dp07 & "," & worker_a(ki-1)
'			end if
'		Case "建設部"  
'			if  stra_dp08="" then
'				stra_dp08= worker_a(ki-1)
'			else  
'				stra_dp08= stra_dp08 & "," & worker_a(ki-1)
'			end if
'		Case "社企"  
'			if  stra_dp09="" then
'				stra_dp09= worker_a(ki-1)
'			else  
'				stra_dp09= stra_dp09 & "," & worker_a(ki-1)
'			end if
'		Case "我家農業"  
'			if  stra_dp10="" then
'				stra_dp10= worker_a(ki-1)
'			else  
'				stra_dp10= stra_dp10 & "," & worker_a(ki-1)
'			end if
'		Case Else     
'	End Select

	Select Case staff_gp_a(ki-1)
		Case "業1" 
			if  stra_dpa1="" then
				stra_dpa1= worker_a(ki-1)
			else  
				stra_dpa1= stra_dpa1 & "," & worker_a(ki-1)
			end if
		Case "業2" 
			if  stra_dpa2="" then
				stra_dpa2= worker_a(ki-1)
			else  
				stra_dpa2= stra_dpa2 & "," & worker_a(ki-1)
			end if
		Case "業3" 
			if  stra_dpa3="" then
				stra_dpa3= worker_a(ki-1)
			else  
				stra_dpa3= stra_dpa3 & "," & worker_a(ki-1)
			end if
		Case "業4" 
			if  stra_dpa4="" then
				stra_dpa4= worker_a(ki-1)
			else  
				stra_dpa4= stra_dpa4 & "," & worker_a(ki-1)
			end if
		Case "業5" 
			if  stra_dpa5="" then
				stra_dpa5= worker_a(ki-1)
			else  
				stra_dpa5= stra_dpa5 & "," & worker_a(ki-1)
			end if
		Case Else     
	End Select
	
next




%>


	