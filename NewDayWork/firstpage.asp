<%@ Language=VBScript CodePage=950 %>
<%
  ' array_worker_crew
  ' �u�@�H���}�Cdaywork.mdb worker_data
  dim workerArr()
  dim cWorkerArr()
  dim eWorkerArr()
  dim staffArr()
  dim staffIdArr()
  dim staffDpArr()
  dim staffGpArr()

  ' �s��Access��Ʈwdaywork.mdb
  DBpath = Server.MapPath("./database/crew.mdb")
  connStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

  '�إ߸�Ʈw�s������
  set conn = Server.CreateObject("ADODB.Connection")

  '�s����Ʈw
  conn.Open connStr

  '�}�Ҹ�ƪ�W��
  tbName = "crew"

  '�إ߸�Ʈw�s������
  set rs=Server.CreateObject("ADODB.Recordset")
  SQLstr = "SELECT * FROM " & tbName &" WHERE hide = false ORDER BY wk_gp_sq ASC"
  rs.open SQLstr, conn, 3, 1

  ' �p�����`��
  nWorker = rs.RecordCount

  ' ���]�}�C�ƥ�
  redim workerArr(Cint(nWorker))
  redim cWorkerArr(Cint(nWorker))
  redim eWorkerArr(Cint(nWorker))
  redim staffArr(Cint(nWorker))
  redim staffIdArr(Cint(nWorker))
  redim staffDpArr(Cint(nWorker))
  redim staffGpArr(Cint(nWorker))

  rs.MoveFirst
  for i = 0 to nWorker-1
    workerArr(i) = rs.Fields.Item("worker")     ' ����W
    cWorkerArr(i) = rs.Fields.Item("wkr_name")  ' ������W
    eWorkerArr(i) = rs.Fields.Item("e_name")    ' �^��W
    staffArr(i) = rs.Fields.Item("e_name")      ' �ʺ�
    staffIdArr(i) = rs.Fields.Item("wkr_id")    ' id
    staffDpArr(i) = rs.Fields.Item("wk_gp")     ' ����
    staffGpArr(i) = rs.Fields.Item("wk_sgp")    ' �s��

    ' ����U�@���O��
    rs.MoveNext
  next

  ' ������ƶ�
  rs.Close
  ' ���]����ܼ�
  set rs = Nothing
  ' ������Ʈw
  conn.Close
  ' ���]�����ܼ�
  set conn = Nothing

  ' ======�����H���r��============
  dp01Str = ""  ' �`�g�z��
  dp02Str = ""  ' �޲z��
  dp03Str = ""  ' ������
  dp04Str = ""  ' �~�ȳ�
  dp05Str = ""  ' �k��+������
  dp06Str = ""  ' �]�ȳ�
  dp07Str = ""  ' ��T+�޲z��
  dp08Str = ""  ' �س]��
  dp09Str = ""  ' ��������|
  dp10Str = ""  ' �ڮa�A�~

  dpA1Str = ""  ' �~1
  dpA2Str = ""  ' �~2
  dpA3Str = ""  ' �~3
  dpA4Str = ""  ' �~4
  dpA5Str = ""  ' �~5

  for i = 1 to nWorker-1
    if InStr(1, staffDpArr(i), "�`�g�z��", 1) > 0 then
        if dp01Str = "" then
          dp01Str = workerArr(i)
        else
          dp01Str = dp01Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "�޲z��", 1) > 0 then
        if dp02Str = "" then
          dp02Str = workerArr(i)
        else
          dp02Str= dp02Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "������", 1) > 0 then
        if dp03Str = "" then
          dp03Str = workerArr(i)
        else
          dp03Str = dp03Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "�~�ȳ�", 1) > 0 then
        if dp04Str = "" then
          dp04Str = workerArr(i)
        else
          dp04Str = dp04Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "�k�ȳ�", 1) > 0 then
        if dp05Str = "" then
          dp05Str = workerArr(i)
        else
          dp05Str = dp05Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "�]�ȳ�", 1) > 0 then
        if dp06Str = "" then
          dp06Str = workerArr(i)
        else
          dp06Str = dp06Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "��T��", 1) > 0 then
        if dp07Str = "" then
          dp07Str = workerArr(i)
        else
          dp07Str = dp07Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "�س]��", 1) > 0 then
        if dp08Str = "" then
          dp08Str = workerArr(i)
        else
          dp08Str = dp08Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "����", 1) > 0 then
        if dp09Str = "" then
          dp09Str = workerArr(i)
        else
          dp09Str = dp09Str & "," & workerArr(i)
        end if
    end if
    if InStr(1, staffDpArr(i), "�ڮa�A�~", 1) > 0 then
        if dp10Str = "" then
          dp10Str = workerArr(i)
        else
          dp10Str = dp10Str & "," & workerArr(i)
        end if
    end if

    Select Case staffGpArr(i)
      Case "�~1"
        if dpA1Str = "" then
          dpA1Str = workerArr(i)
        else
          dpA1Str = dpA1Str & "," & workerArr(i)
        end if
      Case "�~2"
        if dpA2Str = "" then
          dpA2Str = workerArr(i)
        else
          dpA2Str = dpA2Str & "," & workerArr(i)
        end if
      Case "�~3"
        if dpA3Str = "" then
          dpA3Str = workerArr(i)
        else
          dpA3Str = dpA3Str & "," & workerArr(i)
        end if
      Case "�~4"
        if dpA4Str = "" then
          dpA4Str = workerArr(i)
        else
          dpA4Str = dpA4Str & "," & workerArr(i)
        end if
      Case "�~5"
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
  ' �]�wSession�ܼƮ����ɶ�
  worker = Session("worker")
%>

<%
  ' �s��Access��Ʈwdaywork.mdb
  DBpath = Server.MapPath("./database/daywork.mdb")
  connStr ="Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

  ' �إ߸�Ʈw�s������
  set conn = Server.CreateObject("ADODB.Connection")

  ' �s����Ʈw
  conn.Open connStr

  ' �}�Ҹ�ƪ�W��
  tbName = "work_data"

  ' �إ߸�Ʈw�s������
  set rs = Server.CreateObject("ADODB.Recordset")
  SQLstr = "SELECT * FROM " & tbName & " WHERE headline > 5 AND doing_date1 = #" & date() & "# ORDER BY wk_item ASC"
  rs.open SQLstr, conn, 1, 3
  nHeadline = rs.RecordCount  ' ���j�T���ƶq

  dim headlineDate()  ' ���j�T�����
  dim headlineTxt()   ' ���j�T�����e


  if nHeadline = 0 then
    redim headlineDate(1)
    redim headlineTxt(1)
    headlineDate(0) = date()
    headlineTxt(0) = "�L"
  else
    redim headlineDate(nHeadline)
    redim headlineTxt(nHeadline)

    ' �C�X��ƶ���
    rs.MoveFirst
    for i = 0 to nHeadline-1
      headlineDate(i) = rs.Fields.Item("doing_date1")
      headlineTxt(i) = rs.Fields.Item("wk_item")

      ' ����U�@���O��
      rs.MoveNext
      if rs.EOF = True then exit for
    next
  end if

  ' ������ƶ�
  rs.Close
  ' ���]����ܼ�
  set rs = Nothing
  ' ������Ʈw
  conn.Close
  ' ���]�����ܼ�
  set conn = Nothing
%>

<html>
  <head>
    <title>�u�@�޲z�t��</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <link rel="icon" href="../daywork/img/khouse.ico" type="image/vnd.microsoft.icon"/>
    <link rel="stylesheet" href="./css/global.css" type="text/css">
    <style>
      body{
        background: #f0fff0;
        font-family: "�L�n������";
        font-weight: bold;
        

        margin-top: 5px;
      }

      input, td{
        font: 16.66px "�L�n������";
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

        font: 32px "�L�n������";
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
        font: 21.33px "�L�n������";
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
        font: 10pt "�L�n������";
        cursor: hand;
      }

      .btn:hover {
        background: #ffd700;
      }

      .footer-link {
        display: inline-block;
        width: 62px;

        color: #000000 !important;
        font-family: "�L�n������";
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
        font: 16.66px "�L�n������";
        text-decoration: underline;
        cursor: hand;
      }

      .url-small {
        height: 28px;

        color: #830742;
        font: 14.66px "�L�n������";
        text-decoration: underline;
        cursor: hand;
      }

    </style>
  </head>

  <body>
    <div class="center">
      <!-- �лx�Ϥ� -->
      <img src="./img/work_title.jpg">

      <!-- �]���O�}�l -->
      <div class="marquee cell">
        <div>
          <% if nHeadline = 0 then %>
            <span>&nbsp;</span>
          <% else %>
            <span class="red">�T�����i(<%=nHeadline%>��)�G</span>
            <% for i = 0 to nHeadline-1 %>
              <span class="red">&nbsp;&nbsp;<%=i+1%></span>
              �B
              <span class="blue"><%=headlineTxt(i)%>�C</span>
            <% next %>
          <% end if %>
        </div>
      </div>
      <!-- �]���O���� -->

      <div class="content">
        <div class="content-left cell center">
          <div class="title cell center">�w��i�J�ӤH�u�@�޲z�t��!!</div>
          <div class="cell">
            <a class="header-link-primary" href="" target="_blank">�i�G�i��j</a>&nbsp;&nbsp;
            <a class="header-link" href="" target="_blank">�Ȥᤵ��Ω���جP�C��</a>&nbsp;&nbsp;
            <a class="header-link" href="" target="_blank">[����]</a>&nbsp;&nbsp;
            <a class="header-link" href="" target="_blank">[�U��]</a>
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
            <a class="footer-link a00" href="" target="_blank" title="�ΦL�ӽ�">�ΦL</a>
            <a class="footer-link a01" href="./00_01_build.asp" target="_blank" title="�ΦL�ӽ�">�س]��</a>
            <a class="footer-link a02" href="./00_02_sales.asp" target="_blank" title="�ΦL�ӽ�">�~�ȳ�</a>
            <a class="footer-link a03" href="./00_03_manager.asp" target="_blank" title="�ΦL�ӽ�">�޲z��</a>
            <a class="footer-link a04" href="./00_04_finance.asp" target="_blank" title="�ΦL�ӽ�">�]�ȳ�</a>
            <a class="footer-link a05" href="./00_05_law.asp" target="_blank" title="�ΦL�ӽ�">�k�ȳ�</a>
            <a class="footer-link a06" href="./00_06_mis.asp" target="_blank" title="�ΦL�ӽ�">��T��</a>
            <a class="footer-link a07" href="./00_07_golf.asp" target="_blank" title="�ΦL�ӽ�">������</a>
            <a class="footer-link a08" href="./00_08_fundation.asp" target="_blank" title="�ΦL�ӽ�">�� ��</a>
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
            document.write("<tr class='calendar-header1'><th colspan=7>����"+(yr-1911)+"�~"+(month+1)+"��"+date+"��</th>");
            document.write("<tr class='calendar-header2'><th>��</th><th>�@</th><th>�G</th><th>�T</th><th>�|</th><th>��</th><th>��</th></tr>");

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

          <img class="demo-img" src="./img/demo.png" alt="�n�J�i�ܪ�����" onclick="demo()">
        </div>
      </div>

      <%
        ' �s������
        ' �s��Access��Ʈwlinkweb.mdb
        DBpath = Server.MapPath("./database/linkweb.mdb")
        connStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

        ' �إ߸�Ʈw�s������
        set conn = Server.CreateObject("ADODB.Connection")

        ' �s����Ʈw
        conn.Open connStr

        ' �}�Ҹ�ƪ�W��
        tbName = "linkdata"

        ' �إ߸�Ʈw�s������
        set rs = Server.CreateObject("ADODB.Recordset")
        SQLstr="Select * from " & tbName & " order by lk_row asc, lk_col asc"
        rs.Open SQLstr, conn, 3, 1
      %>

      <table class="url-table cell">
        <colgroup span="6"></colgroup>
        <%
          ' �p�����`��
          nLink = rs.RecordCount
          if nLink <> 0 then
            ' ���ܲĤ@�����
            rs.MoveFirst

            prevRow = 0
            ' �C�X��ƶ���
            for i = 1 to nLink
              ' �]�w�ťո�Ƥ��ϬM
              linkId = rs.Fields.Item("lk_id")        ' id
              linkUrl = rs.Fields.Item("lk_url")      ' �s�����}
              linkItem = rs.Fields.Item("lk_item")	  ' �u���D
              linkTitle = rs.Fields.Item("lk_title")	' �y�z
              linkRow = rs.Fields.Item("lk_row")		  ' �C
              linkCol = rs.Fields.Item("lk_col")		  ' ��

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

              ' ����U�@���O��
              rs.MoveNext

              if rs.EOF = True then exit for
            next

            Response.Write("</tr>")
          end if
        %>
      </table>

      <%
        ' ������ƶ�
        rs.Close
        ' ���]����ܼ�
        set rs = Nothing
        ' ������Ʈw
        conn.Close
        ' ���]�����ܼ�
        set conn = Nothing
      %>
    </div>

    <script>
      // �i�ܪO
      function demo() {
        if (confirm("�T�w�i�J�i�ܪ��������I�I")) {
          // �T�w�i�J�i�ܪ�������
          location.href = ""
        }
      }
    </script>
  </body>
</html>