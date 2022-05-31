<%@ Language=VBScript CodePage=950 %>
<%
  ' 設定Session變數消滅時間
  Session.TimeOut = 480

  workerOld = Session("worker")
  if Request.QueryString("fp") = "1" then workerOld = "喬大"

  worker = Request.QueryString("worker")
  Session("worker") = worker
%>

<%
  ' 讀取密碼資料
  Function findCrewPwd(wkr)
    ' 連結Access資料庫crew.mdb
    DBpath = Server.MapPath("./database/crew.mdb")
    connStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

    ' 建立資料庫連結物件
    set conn = Server.CreateObject("ADODB.Connection")

    ' 連結資料庫
    conn.Open connStr

    ' 開啟資料表名稱
    tbName = "crew"

    ' 建立資料庫存取物件
    set rs = Server.CreateObject("ADODB.Recordset")
    SQLstr = "SELECT * FROM " & tbName & " WHERE worker LIKE '"&wkr&"'"
    rs.open SQLstr, conn, 3, 1

    result = rs.RecordCount
    if result = 0 then
      pwd = "0"
    else
      pwd = rs.Fields.Item("wkr_pwd")   ' 密碼
    end if

    ' 關閉資料集
    rs.Close
    ' 重設資料變數
    set rs = Nothing
    ' 關閉資料庫
    conn.Close
    ' 重設物件變數
    set conn = Nothing

    findCrewPwd = pwd
  End Function
%>

<%
  wkr_pwd = Session("wkr_pwd")    ' 讀取密碼
  if Session("wkr_pwd") = "" or IsNull(wkr_pwd) then
    wkr_pwd = Request("wkr_pwd")  ' 讀取密碼
  end if

  chk_str = ""

  if InStr(1, chk_str, worker, 1) > 0 AND InStr(1, chk_str, workerOld, 1) = 0 then
    ' 讀取資料庫密碼
    dbPwd = findCrewPwd(worker)
    if dbPwd = wkr_pwd then
      Session("wkr_pwd") = dbPwd
    else
      url = "./0_login_pwd.asp?worker=" & worker
      Response.Redirect url   ' 轉址到密碼輸入畫面
    end if
  end if
%>

<html>
  <head>
    <title>工作管理系統</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <link rel="icon" href="../daywork/img/khouse.ico" type="image/vnd.microsoft.icon"/>

    <style>
      .title-frame {
        overflow: hidden;
      }
      .main-frame {
        height: calc(100% - 85px);
        overflow: auto;
      }
    </style>
  </head>

  <body>
    <iframe class="title-frame" src="work_title.asp?worker=<%=worker%>" width="100%" height="85"></iframe>
    <iframe class="main-frame" src="wk_Calendar_all.asp?worker=worker" width="100%"></iframe>
  </body>
</html>
