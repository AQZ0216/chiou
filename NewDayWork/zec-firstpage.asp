<%@ Language=VBScript CODEPAGE=950 %>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<%
'headline_no   '���j�T���ƶq
dim headline_txt()      '���j�T�����e
dim headline_date()   '���j�T�����

      ' �s��Access��Ʈwdaywork.mdb
      DBpath_hdl=Server.MapPath("./database/daywork.mdb")
      strCon_hdl="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_hdl
      '�إ߸�Ʈw�s������
      set conDB_hdl= Server.CreateObject("ADODB.Connection")
      '�s����Ʈw	
      conDB_hdl.Open strCon_hdl
      '�}�Ҹ�ƪ�W��
      tb_name_hdl="work_data"
      '�إ߸�Ʈw�s������	
      set rstObj1_hdl=Server.CreateObject("ADODB.Recordset")
      strSQL_show_hdl="Select * from " & tb_name_hdl & " where headline > 5 and doing_date1 = #"& date() &"# order by wk_item asc"   
      rstObj1_hdl.open strSQL_show_hdl,conDB_hdl,1,3
totalput_hdl=rstObj1_hdl.recordcount
headline_no=totalput_hdl
if totalput_hdl=0 then
   redim headline_txt(1)
   redim headline_date(1)
   headline_date(0)=date()
   headline_txt(0)="�L"
else
   redim headline_txt(headline_no)
   redim headline_date(headline_no)
	'�C�X��ƶ���
	rstobj1_hdl.MoveFirst
	for i=1 to totalput_hdl
	     headline_date(i-1)=rstObj1_hdl.fields("doing_date1")
	 	headline_txt(i-1)=rstObj1_hdl.fields("wk_item")
	'����U�@���O��
		rstObj1_hdl.MoveNext
		if rstObj1_hdl.EOF=True then exit for
	next	
end if	
      '������ƶ�
      rstObj1_hdl.Close
      '���]����ܼ�
      set rstObj1_hdl=Nothing
      '������Ʈw
      conDB_hdl.Close
      '���]�����ܼ�
      set conDB_hdl=Nothing
      
str_marquee="�T�����i("& totalput_hdl &"��)�G�@�@"
for zi=1 to headline_no
   str_marquee = str_marquee & zi & "�B" & headline_txt(zi-1) & "�C�@�@"
next
%>
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<title>��j�a���u�@�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="./css/w3-cht.css">
<link rel="stylesheet" href="./css/font-awesome.min.css">
    <link type="text/css" rel="stylesheet" href="./src/css/jscal2.css" />
    <link type="text/css" rel="stylesheet" href="./src/css/border-radius.css" />
    <link type="text/css" rel="stylesheet" href="./src/css/large-spacing.css" />    
    <script src="./src/js/jscal2.js"></script>
    <script src="./src/js/lang/b5.js"></script>
    <script src="./js/w3.js"></script>
<style>
.marquee {
  height: 50px;
  overflow: hidden;
  position: relative;
  background: #fefefe;
  color: #333;
  border: 1px solid #4a4a4a;
  font-family:�L�n������;
}

.marquee span {
  font-family:�L�n������;
  font-size:25px;
  text-overflow: ellipsis; /*�W�X������...�N��*/
  white-space:nowrap;/*�j���r�b�@�椺���*/
  position: absolute;
  width: 100%;
  height: 100%;
  margin: 0;
  line-height: 50px;
  text-align: center;
  -moz-transform: translateX(100%);
  -webkit-transform: translateX(100%);
  transform: translateX(100%);
  -moz-animation: scroll-left 2s linear infinite;
  -webkit-animation: scroll-left 2s linear infinite;
  /*animation: scroll-left 20s linear infinite;*/
  animation: scroll-left 20s linear infinite;
}

@-moz-keyframes scroll-left {
  0% {
      -moz-transform: translateX(100%);
  }
  100% {
      -moz-transform: translateX(-100%);
  }
}

@-webkit-keyframes scroll-left {
  0% {
      -webkit-transform: translateX(100%);
  }
  100% {
      -webkit-transform: translateX(-100%);
  }
}

@keyframes scroll-left {
  0% {
      -moz-transform: translateX(100%);
      -webkit-transform: translateX(100%);
      transform: translateX(100%);
  }
  100% {
      -moz-transform: translateX(-100%);
      -webkit-transform: translateX(-100%);
      transform: translateX(-100%);
  }
}
.marquee span:hover {
    -webkit-animation-play-state: paused;
       -moz-animation-play-state: paused;
        -ms-animation-play-state: paused;
         -o-animation-play-state: paused;
            animation-play-state: paused;
            /*animation-play-state: running;*/
}
</style>
</head>
<body class="vt-container">
<!-- ���Y�}�l -->
<header class="vt-container w3-brown w3-center " style="overflow: hidden;">
  <button class="w3-button w3-brown w3-xxlarge w3-round-xlarge" onclick="location.reload()" title="���㭶��">�i��j�a���j�u�@�޲z����</button>
   <div class="vt-container marquee w3-pale-green w3-border w3-border-brown">
      <span><%=str_marquee%></span>
   </div>
</header>
<!-- ���Y���� -->
<!-- ����@�}�l -->
<div class="vt-container w3-row w3-pale-green w3-border w3-border-brown">
  <div class="w3-col l9 w3-center w3-pale-green">
      <div class="w3-row w3-center " >
         <h3>��j�a���P��</h3>
         <% for i=1 to worker_no %>
   		 	<button class="w3-button w3-large w3-pale-red  w3-border w3-border-red w3-round-large " style="margin:1px;padding:0px;width:100px;" onclick="url_show('./zec-work_month_r1.asp?worker=<%=worker_a(i-1)%>&fp=1')" >
            <span style="font-size:16px;"><%=eworker_a(i-1)%></span><br><span><%=worker_a(i-1)%></span>
            </button>
         <% next %>   
      </div>
  </div>
  <div class="w3-col l3 w3-center w3-pale-green">
      <div class="w3-row w3-center" >
         <h3>���� (<%=date()%>)</h3>
      <div id="calendar-container" ></div>
      	<script type="text/javascript">
      	var flatCalendar=new Calendar({
      		fdow 		:0,						//�C�P�Ĥ@��,0=Sun
      		cont		:"calendar-container",				//�T�w��m�Ӹ� div id
      		selectionType	:Calendar.SEL_SINGLE,				//������Υi�ƿ�
      		//selection	:Calendar.dateToInt(new Date()),
           titleFormat: "%B %Y",
      		showTime	:false,
      		bottomBar	:true,
      		weekNumbers	:true
      	});
      	</script>  
      </div>  
  </div>
</div>
<!-- ����@���� -->
<!-- ����G�}�l -->
<div class="vt-container w3-row l12 w3-center w3-pale-green w3-border w3-border-brown" >
   <button class="w3-button w3-large w3-blue" onclick="url_new('../chopman/zec-z0_login.asp')">�ΦL�ӽ�</button>
   <button class="w3-button w3-large w3-brown" onclick="url_new('./zec-00_01_build.asp')">�س]��</button>
   <button class="w3-button w3-large w3-blue-grey" onclick="url_new('./zec-00_02_sales.asp')">�~�ȳ�</button>
   <button class="w3-button w3-large w3-indigo" onclick="url_new('./zec-00_03_manager.asp')">�޲z��</button>
   <button class="w3-button w3-large w3-pink" onclick="url_new('./zec-00_04_finance.asp')">�]�ȳ�</button>
   <button class="w3-button w3-large w3-purple" onclick="url_new('./zec-00_05_law.asp')">�k�ȳ�</button>
   <button class="w3-button w3-large w3-deep-purple" onclick="url_new('./zec-00_06_mis.asp')">��T��</button>
   <button class="w3-button w3-large w3-red" onclick="url_new('./zec-00_07_golf.asp')">������</button>
   <button class="w3-button w3-large w3-teal" onclick="url_new('./zec-00_08_fundation.asp')">�� ��</button>
</div>
<!-- ����G���� -->
<!-- ����T�}�l -->
<div class="vt-container w3-row w3-border w3-border-brown">
  <div class="w3-col l12 w3-center w3-pale-blue">
      <div class="w3-col w3-center " >
         <h3><button onclick="w3.toggleShow('.linkdata')">�����s��</button></h3>
         <div class="linkdata" >
<%
' �s��Access��Ʈw./database/linkweb.mdb
DBpath=Server.MapPath("./database/linkweb.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'�إ߸�Ʈw�s������
set conDB= Server.CreateObject("ADODB.Connection")
'�s����Ʈw	
conDB.Open strCon
'�}�Ҹ�ƪ�W��
tb_name="linkdata"
'�إ߸�Ʈw�s������	
set rstObj1=Server.CreateObject("ADODB.Recordset")
	strSQL_show="Select * from " & tb_name & " order by lk_row asc, lk_col asc"
rstObj1.open strSQL_show,conDB,3,1
'�p�����`��	
totalput=rstObj1.recordcount
if totalput=0 then
else
      '���ܲĤ@�����
      rstobj1.MoveFirst
      '�C�X��ƶ���
      for i=1 to totalput
      	'�]�w�ťո�Ƥ��ϬM
      p_id=rstObj1.fields("lk_id")		'id	
      p_01=rstObj1.fields("lk_url")		'�s�����}
      p_02=rstObj1.fields("lk_item")		'�u���D
      p_03=rstObj1.fields("lk_title")		'�y�z
      p_04=rstObj1.fields("lk_row")		'�C
      p_05=rstObj1.fields("lk_col")		'��
%>
<button class="w3-button w3-large w3-pale-yellow  w3-border w3-border-brown w3-round-large " style="margin:2px;padding:3px;width:150px;" onclick="url_new('<%=p_01%>')" >
<%=p_02%>
</button>
<%
      '����U�@���O��
      rstObj1.MoveNext
      if rstObj1.EOF=True then exit for
      next
end if
'������ƶ�
rstObj1.Close
'���]����ܼ� 
set rstObj1=Nothing
'������Ʈw 
conDB.Close
'���]�����ܼ� 
set conDB=Nothing
%>
         </div>
      </div>
  </div>
</div>
<!-- ����T���� -->
<!-- �����}�l -->
<!-- #Include file = "./include/zec-footer_r1.inc" -->
<!-- �������� -->
<script language="JavaScript">
    function url_new(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        //window.location.href = pp_url; //�쭶����s
        window.open(pp_url) ; //�}�ҷs����
        return true;
    }   
    function url_show(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        window.location.href = pp_url; //�쭶����s
        //window.open(pp_url) ; //�}�ҷs����
        return true;
    }   
    function zms_show(pp_url){
        var iframe1=document.getElementById("ifrm_milestone");
        iframe1.src=pp_url;
        return true;
    }    
    function zlb_show(pp_url){
        var iframe1=document.getElementById("ifrm_logbook");
        iframe1.src=pp_url;
        return true;
    }
    function zfi_show(pp_url){
        var iframe1=document.getElementById("ifrm_finance");
        iframe1.src=pp_url;
        return true;
    }
</script>

</body>
</html>






