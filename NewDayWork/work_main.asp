<%@ Language=VBScript CodePage=950 %>
<%
  ' �]�wSession�ܼƮ����ɶ�
  Session.TimeOut = 480

  workerOld = Session("worker")
  if Request.QueryString("fp") = "1" then workerOld = "��j"

  worker = Request.QueryString("worker")
  Session("worker") = worker
%>

<%
  ' Ū���K�X���
  Function findCrewPwd(wkr)
    ' �s��Access��Ʈwcrew.mdb
    DBpath = Server.MapPath("./database/crew.mdb")
    connStr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DBpath

    ' �إ߸�Ʈw�s������
    set conn = Server.CreateObject("ADODB.Connection")

    ' �s����Ʈw
    conn.Open connStr

    ' �}�Ҹ�ƪ�W��
    tbName = "crew"

    ' �إ߸�Ʈw�s������
    set rs = Server.CreateObject("ADODB.Recordset")
    SQLstr = "SELECT * FROM " & tbName & " WHERE worker LIKE '"&wkr&"'"
    rs.open SQLstr, conn, 3, 1

    result = rs.RecordCount
    if result = 0 then
      pwd = "0"
    else
      pwd = rs.Fields.Item("wkr_pwd")   ' �K�X
    end if

    ' ������ƶ�
    rs.Close
    ' ���]����ܼ�
    set rs = Nothing
    ' ������Ʈw
    conn.Close
    ' ���]�����ܼ�
    set conn = Nothing

    findCrewPwd = pwd
  End Function
%>

<%
  wkr_pwd = Session("wkr_pwd")    ' Ū���K�X
  if Session("wkr_pwd") = "" or IsNull(wkr_pwd) then
    wkr_pwd = Request("wkr_pwd")  ' Ū���K�X
  end if

  chk_str = ""

  if InStr(1, chk_str, worker, 1) > 0 AND InStr(1, chk_str, workerOld, 1) = 0 then
    ' Ū����Ʈw�K�X
    dbPwd = findCrewPwd(worker)
    if dbPwd = wkr_pwd then
      Session("wkr_pwd") = dbPwd
    else
      url = "./0_login_pwd.asp?worker=" & worker
      Response.Redirect url   ' ��}��K�X��J�e��
    end if
  end if
%>

<html>
  <head>
    <title>�u�@�޲z�t��</title>
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
