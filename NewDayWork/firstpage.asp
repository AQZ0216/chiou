<%@ Language=VBScript CodePage=950 %>
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
  ' 設定Session變數消滅時間
  worker = Session("worker")
%>

<%
  ' 連結Access資料庫daywork.mdb
  DBpath = Server.MapPath("./database/daywork.mdb")
  connStr ="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

  ' 建立資料庫連結物件
  set conn = Server.CreateObject("ADODB.Connection")

  ' 連結資料庫
  conn.Open connStr

  ' 開啟資料表名稱
  tbName = "work_data"

  ' 建立資料庫存取物件
  set rs = Server.CreateObject("ADODB.Recordset")
  SQLstr = "SELECT * FROM " & tbName & " WHERE headline > 5 AND doing_date1 = #" & date() & "# ORDER BY wk_item ASC"
  rs.open SQLstr, conn, 1, 3
  nHeadline = rs.RecordCount  ' 重大訊息數量

  dim headlineDate()  ' 重大訊息日期
  dim headlineTxt()   ' 重大訊息內容


  if nHeadline = 0 then
    redim headlineDate(1)
    redim headlineTxt(1)
    headlineDate(0) = date()
    headlineTxt(0) = "無"
  else
    redim headlineDate(nHeadline)
    redim headlineTxt(nHeadline)

    ' 列出資料項目
    rs.MoveFirst
    for i = 0 to nHeadline-1
      headlineDate(i) = rs.Fields.Item("doing_date1")
      headlineTxt(i) = rs.Fields.Item("wk_item")

      ' 移到下一筆記錄
      rs.MoveNext
      if rs.EOF = True then exit for
    next
  end if

  ' 關閉資料集
  rs.Close
  ' 重設資料變數
  set rs = Nothing
  ' 關閉資料庫
  conn.Close
  ' 重設物件變數
  set conn = Nothing
%>

<html>
  <head>
    <title>工作管理系統</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <link rel="icon" href="../daywork/img/khouse.ico" type="image/vnd.microsoft.icon"/>
    <link rel="stylesheet" href="./css/global.css" type="text/css">
    <style>
      body{
        background: #f0fff0;
        font-family: "微軟正黑體";
        font-weight: bold;
        

        margin-top: 5px;
      }

      input, td{
        font: 16.66px "微軟正黑體";
        cursor: hand;
      }

      a {
        text-decoration: none;
      }
      a:link {
        color: blue;
      }
      a:visited {
        color: blue;
      }
      a:hover {
        color: red;
      }
      a:active {
        color: green;
      }

      .red {
        color: red;
      }

      .blue {
        color: blue;
      }

      .cell {
        border: 1px solid black;
      }

      .marquee {
        width: 800px;
        margin: auto;
        background: #cff3c0;
        white-space: nowrap;
        overflow: hidden;

        font: 32px "微軟正黑體";
        letter-spacing: 7.5px;
      }

      .marquee > div {
        animation: marquee 30s linear 10;
        padding-left: 800px;
        width: max-content;
      }

      @keyframes marquee {
        from {
          transform: translateX(0%);
        }
        to {
          transform: translateX(-100%);
        }
      }

      .content {
        display: flex;
        justify-content: center;
      }

      .content-left {
        width: 594px;
        vertical-align: top;
      }

      .content-right {
        width: 204px;
        vertical-align: middle;
      }

      .title {
        color: blue;
        font: 21.33px "微軟正黑體";
      }

      .header-link-primary {
        background: #ccc;
        color: red !important;
        text-decoration: none;
        font-size: 20px;
        letter-spacing: 4px;
      }

      .header-link {
        background: #ddd;
        color: blue !important;
        text-decoration: none;
      }

      .worker-table {
        margin: auto;

        border-collapse: collapse;
      }

      .worker-table td {
        padding: 0;
      }

      .btn {
        width: 66px;
        height: 30px;
        background: #d3d3d3;
        font: 10pt "微軟正黑體";
        cursor: hand;
      }

      .btn:hover {
        background: #ffd700;
      }

      .footer-link {
        display: inline-block;
        width: 62px;

        color: #000000 !important;
        font-family: "微軟正黑體";
        font-weight: bold;
        text-decoration: none;
      }

      .a00 {
        background: #ff79ff;
      }
      .a01 {
        background: #fcc;
      }
      .a02 {
        background: #fda;
      }
      .a03 {
        background: #ffb;
      }
      .a04 {
        background: #cf9;
      }
      .a05 {
        background: #bfe;
      }
      .a06 {
        background: #9ff;
      }
      .a07 {
        background: #ccf;
      }
      .a08 {
        background: #ffb3ff;
      }

      .calendar {
        margin: auto;

        border-collapse: separate;
        border-spacing: 3px;
      }

      .calendar-header1 {
        background: #ffc080;
        font-size: 16px;
      }

      .calendar-header2 {
        background: #c0ffc0;
        font-size: 14.66px;
      }

      .calendar td {
        padding: 2px;
        font-size: 14.66px;
      }

      .today {
        background: #ffc080;
      }

      .weekend {
        background: #ffc0c0;
      }

      .normal {
        background: #fff;
      }

      .demo-img {
        height: 25px;
        vertical-align: middle;
      }

      .url-table {
        width: 802px;
        margin: auto;

        border-collapse: collapse;
      }

      .url-table td {
        padding: 0;
        border: 1px solid black;
      }

      .url-medium {
        height: 28px;

        color: #830742;
        font: 16.66px "微軟正黑體";
        text-decoration: underline;
        cursor: hand;
      }

      .url-small {
        height: 28px;

        color: #830742;
        font: 14.66px "微軟正黑體";
        text-decoration: underline;
        cursor: hand;
      }

    </style>
  </head>

  <body>
    <div class="center">
      <!-- 標誌圖片 -->
      <img src="./img/work_title.jpg">

      <!-- 跑馬燈開始 -->
      <div class="marquee cell">
        <div>
          <% if nHeadline = 0 then %>
            <span>&nbsp;</span>
          <% else %>
            <span class="red">訊息公告(<%=nHeadline%>筆)：</span>
            <% for i = 0 to nHeadline-1 %>
              <span class="red">&nbsp;&nbsp;<%=i+1%></span>
              、
              <span class="blue"><%=headlineTxt(i)%>。</span>
            <% next %>
          <% end if %>
        </div>
      </div>
      <!-- 跑馬燈結束 -->

      <div class="content">
        <div class="content-left cell center">
          <div class="title cell center">歡迎進入個人工作管理系統!!</div>
          <div class="cell">
            <a class="header-link-primary" href="" target="_blank">【佈告欄】</a>&nbsp;&nbsp;
            <a class="header-link" href="" target="_blank">客戶今日及明日壽星列表</a>&nbsp;&nbsp;
            <a class="header-link" href="" target="_blank">[本月]</a>&nbsp;&nbsp;
            <a class="header-link" href="" target="_blank">[下月]</a>
          </div>

          <!-- toolbar_worker_first_e -->
          <table class="worker-table center">
            <%
              nCol = 9
              nLastCol = nWorker mod nCol

              if nLastCol = 0 then
                nRow = Int(nWorker/nCol)
              else
                nRow = Int(nWorker/nCol) + 1
              end if

            for j = 0 to nRow-1
            %>
              <tr>
                <%
                  for k = 0 to nCol-1
                    i = nCol*j + k
                    if i >= nWorker then
                %>
                      <td class="center">
                        <input class="btn" type="button" title="<%=i%>">
                      </td>
                <%
                    else
                %>
                      <td class="center">
                        <input class="btn" type="button" value="<%=eWorkerArr(i)%>" title="<%=workerArr(i)%>" onclick="parent.location.href='./work_main.asp?worker=<%=workerArr(i)%>&fp=1'">
                      </td>
                <%
                    end if
                  next
                %>
              </tr>
            <%
            next
            %>
          </table>

          <div class="cell center">
            <a class="footer-link a00" href="" target="_blank" title="用印申請">用印</a>
            <a class="footer-link a01" href="./00_01_build.asp" target="_blank" title="用印申請">建設部</a>
            <a class="footer-link a02" href="./00_02_sales.asp" target="_blank" title="用印申請">業務部</a>
            <a class="footer-link a03" href="./00_03_manager.asp" target="_blank" title="用印申請">管理部</a>
            <a class="footer-link a04" href="./00_04_finance.asp" target="_blank" title="用印申請">財務部</a>
            <a class="footer-link a05" href="./00_05_law.asp" target="_blank" title="用印申請">法務部</a>
            <a class="footer-link a06" href="./00_06_mis.asp" target="_blank" title="用印申請">資訊部</a>
            <a class="footer-link a07" href="./00_07_golf.asp" target="_blank" title="用印申請">高爾夫</a>
            <a class="footer-link a08" href="./00_08_fundation.asp" target="_blank" title="用印申請">社 企</a>
          </div>
        </div>

        <div class="content-right cell center">
          <script>
            const time = new Date();
            const yr = time.getFullYear();
            const month = time.getMonth();
            const date = time.getDate();

            const firstday = new Date(yr, month, 1).getDay()

            let days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
            if (((yr % 4 == 0) && (yr % 100 != 0)) || (yr % 400 == 0)) {
              days[1] = 29;
            }

            document.write("<table class='calendar'>");
            document.write("<tr class='calendar-header1'><th colspan=7>民國"+(yr-1911)+"年"+(month+1)+"月"+date+"日</th>");
            document.write("<tr class='calendar-header2'><th>日</th><th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th>六</th></tr>");

            let col = 0;
            for (let i = 0; i < firstday; i++) {
              if (col == 0) {
                document.write("<tr class='right'>");
              }

              document.write("<td>&nbsp;</td>");
              col++;
            }
            for (let i = 1; i <= days[month]; i++) {
              if (col == 0) {
                document.write("<tr class='right'>");
              }

              if (i == date) {
                document.write("<td class='today'>"+i+"</td>");
              } else if ((col == 0) || (col == 6)) {
                document.write("<td class='weekend'>"+i+"</td>");
              } else {
                document.write("<td class='normal'>"+i+"</td>");
              }

              col++;
              if (col == 7) {
                document.write("</tr>");
                col = 0;
              }
            }
            if (col == 7) {
              document.write("</tr>");
            }
            document.write("</table>");
          </script>

          <img class="demo-img" src="./img/demo.png" alt="登入展示版網頁" onclick="demo()">
        </div>
      </div>

      <%
        ' 連結網頁
        ' 連結Access資料庫linkweb.mdb
        DBpath = Server.MapPath("./database/linkweb.mdb")
        connStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

        ' 建立資料庫連結物件
        set conn = Server.CreateObject("ADODB.Connection")

        ' 連結資料庫
        conn.Open connStr

        ' 開啟資料表名稱
        tbName = "linkdata"

        ' 建立資料庫存取物件
        set rs = Server.CreateObject("ADODB.Recordset")
        SQLstr="Select * from " & tbName & " order by lk_row asc, lk_col asc"
        rs.Open SQLstr, conn, 3, 1
      %>

      <table class="url-table cell">
        <colgroup span="6"></colgroup>
        <%
          ' 計算資料總數
          nLink = rs.RecordCount
          if nLink <> 0 then
            ' 移至第一筆資料
            rs.MoveFirst

            prevRow = 0
            ' 列出資料項目
            for i = 1 to nLink
              ' 設定空白資料之反映
              linkId = rs.Fields.Item("lk_id")        ' id
              linkUrl = rs.Fields.Item("lk_url")      ' 連結網址
              linkItem = rs.Fields.Item("lk_item")	  ' 短標題
              linkTitle = rs.Fields.Item("lk_title")	' 描述
              linkRow = rs.Fields.Item("lk_row")		  ' 列
              linkCol = rs.Fields.Item("lk_col")		  ' 欄

              if linkRow <> prevRow then
                if linkRow <> 1 then Response.Write("</tr>")
                
                Response.Write("<tr class='center'>")
                prevRow = linkRow
              end if

              if linkItem = "" or IsNull(linkItem) then
                Response.Write("<td>&nbsp;</td>")
              else
                linkClass = "url-medium"
                if Len(linkItem) > 7 then
                  linkClass = "url-small"
                end if
                %>
                <td class="<%=linkClass%>" title="<%=linkTitle%>"><a href="<%=linkUrl%>" target="_blank"><%=linkItem%></a></td>
                <%
              end if

              ' 移到下一筆記錄
              rs.MoveNext

              if rs.EOF = True then exit for
            next

            Response.Write("</tr>")
          end if
        %>
      </table>

      <%
        ' 關閉資料集
        rs.Close
        ' 重設資料變數
        set rs = Nothing
        ' 關閉資料庫
        conn.Close
        ' 重設物件變數
        set conn = Nothing
      %>
    </div>

    <script>
      // 展示板
      function demo() {
        if (confirm("確定進入展示版本網頁！！")) {
          // 確定進入展示版本網頁
          location.href = ""
        }
      }
    </script>
  </body>
</html>