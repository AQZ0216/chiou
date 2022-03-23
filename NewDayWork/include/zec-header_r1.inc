<div class="w3-sidebar w3-bar-block w3-card w3-animate-left w3-green" style="display:none" id="leftMenu">
  <button onclick="closeLeftMenu()" class="w3-bar-item w3-button w3-large">關閉 &times;</button>
    <!--<a href="#" class="w3-bar-item w3-button">Link 1</a>-->
    <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('./zec-firstpage.asp')" title="回首頁">回首頁</button>
    <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('http://192.168.0.11/chiou/att2000/zec-5_card_query.asp')" title="另開視窗，出勤時間">出勤時間</button>
    <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('../holiday/zec-hd_ps_year_list.asp?wkr_id=<%=pwkr_id%>')" title="另開視窗，休假資料">休假資料</button>
    <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('../customer/cr_wk_con.asp?user=<%=worker%>&pswdck=1')" title="客戶查詢">客戶查詢</button>
     <button class="w3-button" onclick="myAccFunc('psg01')">人員更換(1) ▼</button>
     <div id="psg01" class="w3-hide w3-white w3-card">
       <% for i=1 to 10 %>
       <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('./zec-work_month_r1.asp?worker=<%=worker_a(i-1)%>&fp=1')" title="<%=worker_a(i-1)%>(<%=eworker_a(i-1)%>)"><%=worker_a(i-1)%>(<%=eworker_a(i-1)%>)</button>
       <% next %>
     </div>
     <button class="w3-button" onclick="myAccFunc('psg02')">人員更換(2) ▼</button>
     <div id="psg02" class="w3-hide w3-white w3-card">
       <% for i=11 to 20 %>
       <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('./zec-work_month_r1.asp?worker=<%=worker_a(i-1)%>&fp=1')" title="<%=worker_a(i-1)%>(<%=eworker_a(i-1)%>)"><%=worker_a(i-1)%>(<%=eworker_a(i-1)%>)</button>
       <% next %>
     </div>
     <button class="w3-button" onclick="myAccFunc('psg03')">人員更換(3) ▼</button>
     <div id="psg03" class="w3-dropdown-content w3-bar-block w3-card">
       <% for i=21 to 30 %>
       <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('./zec-work_month_r1.asp?worker=<%=worker_a(i-1)%>&fp=1')" title="<%=worker_a(i-1)%>(<%=eworker_a(i-1)%>)"><%=worker_a(i-1)%>(<%=eworker_a(i-1)%>)</button>
       <% next %>
     </div>
     <button class="w3-button" onclick="myAccFunc('psg04')">人員更換(4) ▼</button>
     <div id="psg04" class="w3-dropdown-content w3-bar-block w3-card">
       <% for i=31 to worker_no %>
       <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('./zec-work_month_r1.asp?worker=<%=worker_a(i-1)%>&fp=1')" title="<%=worker_a(i-1)%>(<%=eworker_a(i-1)%>)"><%=worker_a(i-1)%>(<%=eworker_a(i-1)%>)</button>
       <% next %>
     </div>
</div>

<script language="JavaScript">
    function openLeftMenu() {
      document.getElementById("leftMenu").style.display = "block";
    }

    function closeLeftMenu() {
      document.getElementById("leftMenu").style.display = "none";
    }

    function myAccFunc(ppAccid) {
      var x = document.getElementById(ppAccid);
      if (x.className.indexOf("w3-show") == -1) {
        x.className += " w3-show";
        x.previousElementSibling.className += " w3-green";
      } else { 
        x.className = x.className.replace(" w3-show", "");
        x.previousElementSibling.className = 
        x.previousElementSibling.className.replace(" w3-green", "");
      }
    }
</script>
<!-- 標頭開始 -->
<header class="vt-container w3-brown w3-center" style="height:90px;overflow: hidden;margin:0px;padding:0px;">
    <button class="w3-button w3-teal w3-xlarge w3-left" onclick="openLeftMenu()">&#9776;</button> <!-- 左方功能表 -->
<!--    <button class="w3-button w3-teal w3-xlarge w3-right" onclick="openRightMenu()">&#9776;</button>--> <!-- 右方功能表 -->

  <button class="w3-button w3-brown w3-xlarge w3-round-large" onclick="location.reload()" title="重整頁面" style="padding:4px;margin:4px;">
  【<%=worker%>】工作管理
  </button>
  
   <div class="vt-container w3-center" ><!-- 功能表1 start -->
         <div class="w3-bar w3-teal" ><!-- 功能表gp03 start -->
          <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-firstpage.asp')">回首頁</button>
          <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>')">回日曆表</button>
          <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-work_query.asp?worker=<%=worker%>')">工作查詢</button>
          <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-work_add_r1.asp?datecode=<%=date()%>&worker=<%=worker%>')">工作新增</button>
          <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-wk_pj_list.asp')" title="專案工作">專案工作</button>
          <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-2_admin_main.asp')" title="後台管理">後台管理</button>
       </div> <!-- 功能表gp03 end -->     
   </div><!-- 功能表1 end --> 

</header>
<!-- 標頭結束 -->