<%@ Language=VBScript CODEPAGE=950 %>
<%
	'讀取人員姓名
	worker = Session("worker")
%>
<%
'BASP21.DLL將檔案上載程式
'請先執行安裝RegSvr32 Basp21.dll
'可將表單中的text也抓出塞進陣列，直拉response.write 變數，就可以print出來了
'-------------------------------------------------------------------
%>

<HTML> 
<HEAD>
<Title>上傳檔案功能程式</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
</HEAD>
<BODY>
<center>
<%
dim Upload,A,B,Text,Image,ImgName,ImgPath,RC 
'// Upload為BASP21元件使用之變數 
'// A為將取得的值擁有多少Bytes之變數 
'// B為將取得的值轉為二進位碼的變數 
'// Text為Form內的表單值得變數
'// Image為取得來源圖檔絕對路徑(含檔名)之變數 
'// ImgName為來源圖檔檔名之變數 
'// ImgPath為圖檔欲儲存的路徑及檔名之變數
'// RC為檢查檔案是否上傳之變數 
set Upload = server.CreateObject("BASP21") '// 建立BASP21伺服器物件 
A = request.TotalBytes '// 將取得的值得Bytes值 
B = request.BinaryRead(A) '// 將取得的值轉為二進位碼
Text= Upload.Form(B,"text") '// 取得Form值並轉為二進位碼    '工作wk_id
p_item= Upload.Form(B,"item") '// 取得Form值並轉為二進位碼  '檔案說明
Image = Upload.FormFileName(B,"image") '// 取得圖檔絕對路徑(含檔名)     '檔案名稱
ImgName = mid(Image,InStrRev(Image,"\")+1) '// 取得圖檔名稱(含副檔名) \
file= text &"_"& ImgName     '檔案名稱改為 wk_id+原檔名

response.Write "ImgName="& ImgName &"<br>"   '我寫
response.Write "wk_id="& text &"<br>"  '我寫

'ImgPath = server.MapPath("addfile") & "\" & ImgName '//你欲儲存檔案到哪裡的路徑及檔名
ImgPath = server.MapPath("file_att") & "\" & file '//你欲儲存檔案到哪裡的路徑及檔名

'// 檢查伺服器指定的路徑是否有相同的檔案
if Upload.FileCheck(ImgPath) >= 0 then 
     'set Upload = nothing '//清空物件[註：務必清空，否則秀出來的網頁會變成二進位碼。]
      'response.Write("") &vbcrlf 
      'response.End
      Response.Write "伺服器中有相同檔名 <br>"
      old_file=1
else '// 若檔案不存在
      'RC = Upload.FormSaveAs(B,"image",ImgPath) '//上傳檔案從Form的檔案表單image到伺服器的ImgPath
      '// 檢查檔案上傳成功與否
      Response.Write "伺服器中沒有相同檔名 <br>"
      old_file=0
end if 

'response.Write "ImgPath="& ImgPath &"<br>"   '我寫
'response.end
'============限制副檔名=====================20111118
str_except_file="avi、mpg、mlv、mpe、mpeg、asf、wmv、.rm、rmvb"   '例外副檔名
'file_ext=right(ImgName,InStrRev(ImgName,".",-1,1)-1)
file_ext=right(ImgName,3)
'response.write "副檔名："& file_ext &"<br>"
if instr(1,str_except_file,file_ext,1)=0 then
   RC = Upload.FormSaveAs(B,"image",ImgPath)      '//上傳檔案從Form的檔案表單image到伺服器的ImgPath
else
   RC=0
end if
'============限制副檔名=====================20111118
'RC = Upload.FormSaveAs(B,"image",ImgPath)      '//上傳檔案從Form的檔案表單image到伺服器的ImgPath

set Upload = nothing

if RC > 0 then
      Response.Write "[ "&RC&" ] byte上傳成功 .<br>"
      Response.Write "檔案上傳成功 .<br>"
      '============= 將上傳檔案資料輸入資料庫中 ==================
      p_wk_id=text     'wk_id
      p_fl_name=file   '檔案名稱
      p_fl_size=RC      '檔案大小
      p_fl_date=date() '建檔日期
      p_fl_item=p_item '檔案說明
      p_fl_inputer=worker '上傳檔案人員

    if old_file=0 then
         p_fl_history=now()&"〔"&worker&"〕上傳檔案。"
         ' 連結Access資料庫attach_file.mdb
         DBpath_fl=Server.MapPath("./database/attach_file.mdb")
         strCon_fl="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fl
         '建立資料庫連結物件
         set conDB_fl= Server.CreateObject("ADODB.Connection")
         '連結資料庫	
         conDB_fl.Open strCon_fl
         '開啟資料表名稱
         tb_name_fl="file_data"
         '新增資料之SQL指令字串
         strSQL_add_fl="Insert into "&tb_name_fl&" ("
         strSQL_add_fl=strSQL_add_fl & "wk_id,fl_name,fl_size,fl_date,fl_item,fl_inputer,fl_history) values ('"
         strSQL_add_fl=strSQL_add_fl & p_wk_id &"','"
         strSQL_add_fl=strSQL_add_fl & p_fl_name &"','"
         strSQL_add_fl=strSQL_add_fl & p_fl_size &"',#"
         strSQL_add_fl=strSQL_add_fl & p_fl_date &"#,'"
         strSQL_add_fl=strSQL_add_fl & p_fl_item &"','"
         strSQL_add_fl=strSQL_add_fl & p_fl_inputer &"','"
         strSQL_add_fl=strSQL_add_fl & p_fl_history &"')"
         '執行SQL指令
         conDB_fl.Execute strSQL_add_fl
         '關閉資料庫
         conDB_fl.Close
         '重設物件變數
         set conDB_fl=Nothing
   else
         p_fl_history=now()&"〔"&worker&"〕上傳檔案取代原檔案。"
         ' 連結Access資料庫attach_file.mdb
         DBpath_fl=Server.MapPath("./database/attach_file.mdb")
         strCon_fl="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fl
         '建立資料庫連結物件
         set conDB_fl= Server.CreateObject("ADODB.Connection")
         '連結資料庫	
         conDB_fl.Open strCon_fl
         '開啟資料表名稱
         tb_name_fl="file_data"
         '建立資料庫存取物件	
         set rstObj1_fl=Server.CreateObject("ADODB.Recordset")
         strSQL_show_fl="Select * from " & tb_name_fl & " where wk_id="& p_wk_id & " and fl_name like '"& p_fl_name &"' order by fl_name asc"
         rstObj1_fl.open strSQL_show_fl,conDB_fl,1,3
         rstObj1_fl.fields("fl_size")=p_fl_size
         rstObj1_fl.fields("fl_date")=p_fl_date
         rstObj1_fl.fields("fl_item")=p_fl_item
         rstObj1_fl.fields("fl_inputer")=p_fl_inputer
         rstObj1_fl.fields("fl_history")=rstObj1_fl.fields("fl_history") & chr(13) & p_fl_history
         rstObj1_fl.UpdateBatch
         '關閉資料集
         rstObj1_fl.Close
         '重設資料變數 
         set rstObj1_fl=Nothing
         '關閉資料庫
         conDB_fl.Close
         '重設物件變數
         set conDB_fl=Nothing
   end if
      '============= 將上傳檔案資料輸入資料庫中 ==================
      set Upload = nothing

   myURL="wk_show.asp?wk_id="& text
   Response.Redirect (myURL)
else
      set Upload = nothing '//清空物件[註：務必清空，否則秀出來的網頁會變成二進位碼。]
      Response.Write "[ "&RC&" ] byte上傳 .<br>"
      'response.Write("") &vbcrlf
      'response.End
      'Response.Write "檔案上傳失敗 !<br>"
%>
   <%=ImgName%> 檔案上傳失敗 !<br>
   副檔名為"avi、mpg、mlv、mpe、mpeg、asf、wmv、rm、rmvb"，無法上傳。
<hr>
<a href="wk_show.asp?wk_id=<%=text%>" target="_self">回工作頁面</a>

<%
end if
%>

</center>
</BODY>
</HTML>