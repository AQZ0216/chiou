<%
  ' 地點陣列misc_data.mdb place_table
  dim placeArray()

  ' 連結Access資料庫misc_data.mdb
  DBpath = Server.MapPath("./misc_data/misc_data.mdb")
  connStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

  ' 建立資料庫連結物件
  set conn = Server.CreateObject("ADODB.Connection")

  ' 連結資料庫
  conn.Open connStr

  ' 開啟資料表名稱
  tbName = "place_table"

  ' 建立資料庫存取物件
  set rs = Server.CreateObject("ADODB.Recordset")
  SQLstr =" SELECT * FROM " & tbName
  rs.open SQLstr, conn, 3, 1

  ' 計算資料總數
  nPlace = rs.RecordCount

  ' 重設陣列數目
  redim placeArray(Cint(nPlace))

  rs.MoveFirst
  for i = 0 to nPlace-1
    placeArray(i) = rs.Fields.Item("place")
    ' 移到下一筆記錄
    rs.MoveNext
  next

  ' 關閉資料集
  rs.Close
  ' 重設資料變數
  set rs = Nothing
  ' 關閉資料庫
  conn.Close
  ' 重設物件變數
  set conn = Nothing
%>


