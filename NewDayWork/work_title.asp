<%@ Language=VBScript CodePage=950 %>
<%
  ' 讀取人員姓名
  worker = Session("worker")
%>

<%
  ' array_worker_crew
  ' 工作人員陣列daywork.mdb worker_data
  dim workerArr()
  dim cWorkerArr()
  dim eWorkerArr()
  dim staffArr()
  dim staffIdArr()
  dim staffDpArr()
  dim staffGpArr()

  ' 連結Access資料庫daywork.mdb
  DBpath = Server.MapPath("./database/crew.mdb")
  connStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

  '建立資料庫連結物件
  set conn = Server.CreateObject("ADODB.Connection")

  '連結資料庫
  conn.Open connStr

  '開啟資料表名稱
  tbName = "crew"

  '建立資料庫存取物件
  set rs=Server.CreateObject("ADODB.Recordset")
  SQLstr = "SELECT * FROM " & tbName &" WHERE hide = false ORDER BY wk_gp_sq ASC"
  rs.open SQLstr, conn, 3, 1

  ' 計算資料總數
  nWorker = rs.RecordCount

  ' 重設陣列數目
  redim workerArr(Cint(nWorker))
  redim cWorkerArr(Cint(nWorker))
  redim eWorkerArr(Cint(nWorker))
  redim staffArr(Cint(nWorker))
  redim staffIdArr(Cint(nWorker))
  redim staffDpArr(Cint(nWorker))
  redim staffGpArr(Cint(nWorker))

  rs.MoveFirst
  for i = 0 to nWorker-1
    workerArr(i) = rs.Fields.Item("worker")     ' 中文名
    cWorkerArr(i) = rs.Fields.Item("wkr_name")  ' 全中文名
    eWorkerArr(i) = rs.Fields.Item("e_name")    ' 英文名
    staffArr(i) = rs.Fields.Item("e_name")      ' 暱稱
    staffIdArr(i) = rs.Fields.Item("wkr_id")    ' id
    staffDpArr(i) = rs.Fields.Item("wk_gp")     ' 部門
    staffGpArr(i) = rs.Fields.Item("wk_sgp")    ' 群組

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

  ' ======部門人員字串============
  dp01Str = ""  ' 總經理室
  dp02Str = ""  ' 管理部
  dp03Str = ""  ' 企劃部
  dp04Str = ""  ' 業務部
  dp05Str = ""  ' 法務+企劃部
  dp06Str = ""  ' 財務部
  dp07Str = ""  ' 資訊+管理部
  dp08Str = ""  ' 建設部
  dp09Str = ""  ' 社企基金會
  dp10Str = ""  ' 我家農業

  dpA1Str = ""  ' 業1
  dpA2Str = ""  ' 業2
  dpA3Str = ""  ' 業3
  dpA4Str = ""  ' 業4
  dpA5Str = ""  ' 業5

  for i = 1 to nWorker-1
    if InStr(1, staffDpArr(i), "總經理室", 1) > 0 then
        if dp01Str = "" then
          dp01Str = workerArr(i)
        else
          dp01Str = dp01Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "管理部", 1) > 0 then
        if dp02Str = "" then
          dp02Str = workerArr(i)
        else
          dp02Str= dp02Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "企劃部", 1) > 0 then
        if dp03Str = "" then
          dp03Str = workerArr(i)
        else
          dp03Str = dp03Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "業務部", 1) > 0 then
        if dp04Str = "" then
          dp04Str = workerArr(i)
        else
          dp04Str = dp04Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "法務部", 1) > 0 then
        if dp05Str = "" then
          dp05Str = workerArr(i)
        else
          dp05Str = dp05Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "財務部", 1) > 0 then
        if dp06Str = "" then
          dp06Str = workerArr(i)
        else
          dp06Str = dp06Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "資訊部", 1) > 0 then
        if dp07Str = "" then
          dp07Str = workerArr(i)
        else
          dp07Str = dp07Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "建設部", 1) > 0 then
        if dp08Str = "" then
          dp08Str = workerArr(i)
        else
          dp08Str = dp08Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "社企", 1) > 0 then
        if dp09Str = "" then
          dp09Str = workerArr(i)
        else
          dp09Str = dp09Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "我家農業", 1) > 0 then
        if dp10Str = "" then
          dp10Str = workerArr(i)
        else
          dp10Str = dp10Str & "," & workerArr(i)
        end if
    end if

    Select Case staffGpArr(i)
      Case "業1"
        if dpA1Str = "" then
          dpA1Str = workerArr(i)
        else
          dpA1Str = dpA1Str & "," & workerArr(i)
        end if
      Case "業2"
        if dpA2Str = "" then
          dpA2Str = workerArr(i)
        else
          dpA2Str = dpA2Str & "," & workerArr(i)
        end if
      Case "業3"
        if dpA3Str = "" then
          dpA3Str = workerArr(i)
        else
          dpA3Str = dpA3Str & "," & workerArr(i)
        end if
      Case "業4"
        if dpA4Str = "" then
          dpA4Str = workerArr(i)
        else
          dpA4Str = dpA4Str & "," & workerArr(i)
        end if
      Case "業5"
        if dpA5Str="" then
          dpA5Str = workerArr(i)
        else
          dpA5Str = dpA5Str & "," & workerArr(i)
        end if
      Case Else
    End Select
  next
%>

<%
  for i = 0 to nWorker-1
    if worker = workerArr(i) then
      pwkr_id = staffIdArr(i)
    end if
  next	 
%>

<html>
  <head>
    <title>工作管理系統</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <link rel="stylesheet" href="./css/global.css" type="text/css">

    <style>
      body {
        background: #F0FFF0;
        font-family: "微軟正黑體";
        
        margin-top: 10px;
      }

      input, select {
        font: 16px "微軟正黑體";
      }

      td {
        font-family: "微軟正黑體";
      }

      .homepage {
        width: 90px;
        cursor: hand;
        background: #d3d3d3;
      }

      .homepage:hover {
        background: #ffd700;
      }

      .title {
        color:blue;
        background: #eee;
        font-size: 16pt;
        font-weight: bold;
        letter-spacing: 2pt;
      }

      .icon {
        height: 30px;
        vertical-align: middle;
        cursor: hand;
      }

      .btn {
        width:90px;
        background: #d3d3d3;
        cursor: hand;
        padding: 0;
      }

      .btn:hover {
        background: #ffd700;
      }
    </style>

    <style type="text/css">

    </style>
  </head>

  <body>
    <!-- 標題開始 -->
    <div class="center noBorder">
      <div class="center noPadding">
        <input class="homepage" type="button" value="回首頁" onclick="parent.location.href='firstpage.asp'">
        &nbsp;&nbsp;
        <span class="title">【<%=worker%>】工作管理</span>
        &nbsp;&nbsp;
        <select id="worker" onchange="changeWorker()">
          <option value="" selected>人員更換</option>
          <%
            for i = 0 to nWorker-1
              Response.Write("<option value='"&workerArr(i)&"'>" & workerArr(i) & "</option>")
            next
          %>
        </select>
        &nbsp;&nbsp;
        <img class="icon" src="./img/clock.png" title="查詢出勤時間" onclick="querytime()">
      </div>

      <!-- toolbar_work_title -->
      <div class="center noBorder noPadding">
        <input class="btn center" type="button" value="新增工作" onclick="parent.frames[1].location.href='./wk_add.asp'"></input>
        <input class="btn center" type="button" value="執行工作" onclick="parent.frames[1].location.href='./wk_lst_doing.asp'"></input>
        <input class="btn center" type="button" value="專案工作" onclick="parent.frames[1].location.href='./wk_pj_list.asp'"></input>
        <input class="btn center" type="button" value="預計工作" onclick="parent.frames[1].location.href='./wk_lst_undo.asp'"></input>
        <input class="btn center" type="button" value="完成工作" onclick="parent.frames[1].location.href='./wk_lst_done.asp'"></input>
        <input class="btn center" type="button" value="工作查詢" onclick="parent.frames[1].location.href='./wk_query.asp'"></input>
        <input class="btn center" type="button" value="未完成日曆" onclick="parent.frames[1].location.href='./wk_calendar_all.asp'"></input>
        <input class="btn center" type="button" value="完成日曆" onclick="parent.frames[1].location.href='./wk_calendar_alldone.asp'"></input>
        <input class="btn center" type="button" value="後台管理" onclick="parent.frames[1].location.href='./2_admin_main.asp'"></input>
      </div>
    </div>
    <!-- 標題結束 -->

    <script>
      // 查詢刷卡時間
      function querytime() {
        if (confirm("確定查詢刷卡時間！！。")) {
          // 確定編輯
          // window.open ""
        }
      }

      // 登入選擇人員
      function changeWorker() {
        const worker = document.getElementById("worker").value
        if (worker == "") {
          alert("請選擇人員！！。")
        } else {
          parent.location.href = "./work_main.asp?worker=" + worker
        }
      }
    </script>
  </body>
</html>
