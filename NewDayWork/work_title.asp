<%@ Language=VBScript CodePage=950 %>
<%
  ' Ū���H���m�W
  worker = Session("worker")
%>

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
  for i = 0 to nWorker-1
    if worker = workerArr(i) then
      pwkr_id = staffIdArr(i)
    end if
  next	 
%>

<html>
  <head>
    <title>�u�@�޲z�t��</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <link rel="stylesheet" href="./css/global.css" type="text/css">

    <style>
      body {
        background: #F0FFF0;
        font-family: "�L�n������";
        
        margin-top: 10px;
      }

      input, select {
        font: 16px "�L�n������";
      }

      td {
        font-family: "�L�n������";
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
    <!-- ���D�}�l -->
    <div class="center noBorder">
      <div class="center noPadding">
        <input class="homepage" type="button" value="�^����" onclick="parent.location.href='firstpage.asp'">
        &nbsp;&nbsp;
        <span class="title">�i<%=worker%>�j�u�@�޲z</span>
        &nbsp;&nbsp;
        <select id="worker" onchange="changeWorker()">
          <option value="" selected>�H����</option>
          <%
            for i = 0 to nWorker-1
              Response.Write("<option value='"&workerArr(i)&"'>" & workerArr(i) & "</option>")
            next
          %>
        </select>
        &nbsp;&nbsp;
        <img class="icon" src="./img/clock.png" title="�d�ߥX�Ԯɶ�" onclick="querytime()">
      </div>

      <!-- toolbar_work_title -->
      <div class="center noBorder noPadding">
        <input class="btn center" type="button" value="�s�W�u�@" onclick="parent.frames[1].location.href='./wk_add.asp'"></input>
        <input class="btn center" type="button" value="����u�@" onclick="parent.frames[1].location.href='./wk_lst_doing.asp'"></input>
        <input class="btn center" type="button" value="�M�פu�@" onclick="parent.frames[1].location.href='./wk_pj_list.asp'"></input>
        <input class="btn center" type="button" value="�w�p�u�@" onclick="parent.frames[1].location.href='./wk_lst_undo.asp'"></input>
        <input class="btn center" type="button" value="�����u�@" onclick="parent.frames[1].location.href='./wk_lst_done.asp'"></input>
        <input class="btn center" type="button" value="�u�@�d��" onclick="parent.frames[1].location.href='./wk_query.asp'"></input>
        <input class="btn center" type="button" value="���������" onclick="parent.frames[1].location.href='./wk_calendar_all.asp'"></input>
        <input class="btn center" type="button" value="�������" onclick="parent.frames[1].location.href='./wk_calendar_alldone.asp'"></input>
        <input class="btn center" type="button" value="��x�޲z" onclick="parent.frames[1].location.href='./2_admin_main.asp'"></input>
      </div>
    </div>
    <!-- ���D���� -->

    <script>
      // �d�ߨ�d�ɶ�
      function querytime() {
        if (confirm("�T�w�d�ߨ�d�ɶ��I�I�C")) {
          // �T�w�s��
          // window.open ""
        }
      }

      // �n�J��ܤH��
      function changeWorker() {
        const worker = document.getElementById("worker").value
        if (worker == "") {
          alert("�п�ܤH���I�I�C")
        } else {
          parent.location.href = "./work_main.asp?worker=" + worker
        }
      }
    </script>
  </body>
</html>
