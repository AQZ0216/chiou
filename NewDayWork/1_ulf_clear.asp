<%@ Language=VBScript CODEPAGE=950 %>
<%
'函數檢查工作是否存在
function exist_wkid(pwk_id)
      ' 連結Access資料庫daywork.mdb
      DBpath_fe=Server.MapPath("./database/daywork.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '建立資料庫連結物件
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '連結資料庫	
      conDB_fe.Open strCon_fe
      '開啟資料表名稱
      tb_name_fe="work_data"
      '建立資料庫存取物件	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '關閉資料集
      rstObj1_fe.Close
      '重設資料變數
      set rstObj1_fe=Nothing
      '關閉資料庫 
      conDB_fe.Close
      '重設物件變數
      set conDB_fe=Nothing
      exist_wkid=totalput_fe
end function
%>
<%
' ------------------------------------------
' 將檔案從file_del目錄中刪除
Sub Delfile(strFile)
   'strFile 檔案名稱
'   strDir1=Server.MapPath("./file_att")    '原 用虛擬路徑取得檔案位置
   strDir2=Server.MapPath("./file_del")   '新 用虛擬路徑取得檔案位置
   response.write strFile &"<br>"
'   response.end
      '宣告物件objFSO objInStream及變數intCount strFileContent strInLine
	Dim objFSO, objInStream, intCount, strFileContent, strInLine
	'設定檔案存取物件
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    '刪除舊目錄中之檔案	
	if  (objFSO.FileExists(strDir2 & "\" & strFile)) then
		objFSO.DeleteFile(strDir2 & "\" & strFile)
	else
	end if
	'將檔案移至新目錄
'	Set MyFile = objFSO.GetFile(strDir1 & "\" & strFile)
'	MyFile.Move Server.MapPath("./file_del")& "\"
    Set objFSO = Nothing
    response.write "<hr>"
end sub 
' ------------------------------------------
%>
<%
	'讀取人員姓名
	worker = Session("worker")
	fl_id=Request("fl_id")
%>
<%
'附加檔案列表
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/attach_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="file_data"

'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where fl_id =" & fl_id &" and del_ok = true"
rstObj1.open strSQL_show,conDB,3,3
totalput=rstObj1.recordcount
if totalput=0 then
   del_ok=0
else
   del_ok=1
	'列出資料項目
	rstobj1.MoveFirst
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")      '檔案名稱
         Delfile pfl_name                   '刪除檔案位置
         rstObj1.UpdateBatch
end if
'關閉資料集
rstObj1.Close
'重設資料變數
set rstObj1=Nothing

if del_ok=1 then
   '==============刪除資料============================
   '刪除資料之SQL指令字串
   strSQL_del="Delete from " & tb_name & " where fl_id =" & fl_id
   '執行SQL指令
   conDB.Execute strSQL_del
   '==============刪除資料============================
end if

'關閉資料庫 
conDB.Close
'重設物件變數
set conDB=Nothing

'response.write "檔案刪除完成"
myURL="1_ulf_del_list.asp"
Response.Redirect (myURL)

'if exist_wkid(pwk_id)=1 then
'   'response.write "檔案刪除完成"
'   myURL="wk_show.asp?wk_id="&pwk_id
'   Response.Redirect (myURL)
'else
'   'response.write "檔案刪除完成"
'   myURL="1_ulf_list.asp"
'   Response.Redirect (myURL)
'end if
%>

<HTML> 
<HEAD>
<Title>上傳檔案功能程式</Title>
<META http-equiv="Content-Type" >
<META name="Generator" >
</HEAD>
<BODY>
<center>
檔案刪除完成!!
<hr>
<a href="wk_show.asp?wk_id=<%=pwk_id%>" target="_self">回工作頁面</a>
</center>
</BODY>
</HTML>