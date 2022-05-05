<%
  ' 事件陣列misc_data.mdb thing_table
  dim thingArray()

  ' 連結Access資料庫misc_data.mdb
  DBpath = Server.MapPath("./misc_data/misc_data.mdb")
  connStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

  ' 建立資料庫連結物件
  set conn = Server.CreateObject("ADODB.Connection")

  ' 連結資料庫
  conn.Open connStr

  ' 開啟資料表名稱
  tbName = "thing_table"

  ' 建立資料庫存取物件
  set rs=Server.CreateObject("ADODB.Recordset")
  SQLstr = "SELECT * FROM " & tbName
  rs.open SQLstr, conn, 3, 1

  ' 計算資料總數
  nThing = rs.RecordCount

  ' 重設陣列數目
  redim thingArray(Cint(nThing))

  rs.MoveFirst
  for i = 0 to nThing-1
    thingArray(i) = rs.Fields.Item("thing")
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


