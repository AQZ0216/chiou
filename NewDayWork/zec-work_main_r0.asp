<% @codepage=950%>
<%
   '設定Session變數消滅時間
   Session.TimeOut=480
'Session("worker")=Request("worker")
worker_old = Session("worker")
if request("fp")="1" then worker_old="喬大"
'if worker_old="" or isnull(worker_old) then worker_old="喬大"
worker=Request("worker")
Session("worker")=worker
%>
<%
'讀取密碼資料
Function find_crewpwd(p_wkr)
      ' 連結Access資料庫../holiday/database/crew.mdb
      DBpath_p=Server.MapPath("../holiday/database/crew.mdb")
      strCon_p="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_p
      '建立資料庫連結物件
      set conDB_p= Server.CreateObject("ADODB.Connection")
      '連結資料庫	
      conDB_p.Open strCon_p
      '開啟資料表名稱
      tb_name_p="crew"
      '建立資料庫存取物件	
      set rstObj1_p=Server.CreateObject("ADODB.Recordset")
      strSQL_show_p="Select * from " & tb_name_p & " where worker like '" & p_wkr &"'"
      rstObj1_p.open strSQL_show_p,conDB_p,3,1
      totalp=rstObj1_p.recordcount
      if totalp=0 then
         p_pwd="0"
      else
         p_pwd=rstObj1_p.fields("wkr_pwd")		'密碼
      end if
      '關閉資料集
      rstObj1_p.Close
      '重設資料變數 
      set rstObj1_p=Nothing
      '關閉資料庫
      conDB_p.Close
      '重設物件變數 
      set conDB_p=Nothing
      find_crewpwd=p_pwd
End Function
%>
<%
'---------------------------------------------------------------------------
wkr_pwd=session("wkr_pwd") '讀取密碼
if session("wkr_pwd")="" or isnull(wkr_pwd) then
      wkr_pwd=request("wkr_pwd") '讀取密碼
end if
'chk_str="郭總、國賢、國哲、美慧、惠娟"
chk_str=""
'if instr(1,chk_str,worker,1)>0 and instr(1,chk_str,worker_old,1)=0 then
if instr(1,chk_str,worker,1)>0 then
   if instr(1,chk_str,worker_old,1)>0 then
   'response.write "worker_old="&worker_old
   'response.end
   else
         '讀取資料庫密碼
      '   response.write "worker="&worker&"<br>"
         db_pwd=find_crewpwd(worker)
         if db_pwd=wkr_pwd then
           session("wkr_pwd")=db_pwd
         else
            str_url="./0_login_pwd.asp?worker="&worker
            response.redirect(str_url)      '轉址到密碼輸入畫面
         end if
      end if
end if
'---------------------------------------------------------------------------
%>
<%
'休假資料
function hd_man(p_hdate)
   pstr_hdman =""
    ' 連結Access資料庫holiday.mdb
    DBpath_fh=Server.MapPath("../holiday/database/holiday.mdb")
    strCon_fh="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fh
    '建立資料庫連結物件
    set conDB_fh= Server.CreateObject("ADODB.Connection")
    '連結資料庫	
    conDB_fh.Open strCon_fh
    '開啟資料表名稱
    tb_name_fh="休假明細"
	'建立資料庫存取物件
	set rstObj1_fh=Server.CreateObject("ADODB.Recordset")
	strSQL_show_fh="Select * from " & tb_name_fh & " where 休假日 = #"& p_hdate &"# order by 假別id asc "
	rstObj1_fh.open strSQL_show_fh,conDB_fh,3,1
	totalput_fh=rstObj1_fh.recordcount
if not rstObj1_fh.EOF then
	rstObj1_fh.Movefirst
	for i = 1 to totalput_fh
		hd_id=rstObj1_fh.fields("hd_id")
		icon_id=rstObj1_fh.fields("假別id")
		hd_hrs=rstObj1_fh.fields("休假時數")
		hd_check=rstObj1_fh.fields("確認")
		hd_man=rstObj1_fh.fields("員工姓名")'員工姓名
		hd_img=left(rstObj1_fh.fields("假別名稱"),1)
		hd_cname=right(rstObj1_fh.fields("假別名稱"),len(rstObj1_fh.fields("假別名稱"))-1)
		'決定假別顏色
		select case icon_id
		   Case 1  f_color = "#000000"    '○：公休。
		   Case 2  f_color = "#000000"    '▲：事假。
		   Case 3  f_color = "#000000"    '⊕：病假。
		   Case 4  f_color = "#000000"    '㊣：公假。
		   Case 5  f_color = "#000000"    '◆：喪假。
		   Case 6  f_color = "#000000"    '△：年假。
		   Case 7  f_color = "#000000"    '■：特休。
		   Case 8  f_color = "#000000"    '★：產假。
		   Case 9  f_color = "#000000"    '◎：婚假。
		   Case 15  f_color = "#000000"   '※：未打卡。
		   Case 16  f_color = "#000000"   '▽：排班。
		   Case 17  f_color = "#000000"    '＠：產檢假。
		   Case 18  f_color = "#000000"    '＠：陪產假。
		   Case 19  f_color = "#000000"    '♀：育嬰假。
		   Case Else   f_color = "#000000"
		End Select
		if icon_id=1 or icon_id=15 then
		    if icon_id=15 then
		       pstr_hdman = pstr_hdman & "<font style='font-size:15px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man & "</font><br>"
   	           end if
		else
		  pstr_hdman = pstr_hdman & "<font style='font-size:15px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man &"("& hd_hrs&")&nbsp;</font><br>"
		end if
		    'Response.Write "</font><br>"
		rstObj1_fh.MoveNext
		if rstObj1_fh.EOF=true then exit for
	next
else
end if
	'關閉資料集
	rstObj1_fh.Close
	'重設資料變數 
	set rstObj1_fh=Nothing
    '關閉資料庫
    conDB_fh.Close
    '重設物件變數 
    set conDB_fh=Nothing
  hd_man=pstr_hdman
end function
%>
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<title>【<%=worker%>】工作管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="./css/w3-cht.css">
<style>

</style>
</head>
<body class="vt-container">
<!--標頭-->
<div class="header w3-brown w3-center" style="overflow: hidden;">
  <button class="w3-button w3-brown w3-xlarge w3-round-large" onclick="location.reload()" title="重整頁面" style="padding:4px;margin:4px;">
  【<%=worker%>】工作管理
  </button>
</div>
<div class="w3-pale-blue w3-center" >
    <div class="w3-bar w3-green "><!-- 功能表1 start -->
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('./zec-firstpage.asp')" title="回首頁">回首頁</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('http://192.168.0.11/chiou/att2000/5_card_query.asp')" title="另開視窗，出勤時間">出勤時間</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('../holiday/hd_ps_year_list.asp?wkr_id=<%=pwkr_id%>')" title="另開視窗，休假資料">休假資料</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('http://1.34.48.166:90/firstpage.asp')" title="另開視窗，球場日曆">球場日曆</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('http://60.251.159.62:6980/build/daywork/firstpage.asp?paswd=28283939')" title="另開視窗，建設部">建設部</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_new('../customer/cr_wk_con.asp?user=<%=worker%>&pswdck=1')" title="客戶查詢">客戶查詢</button>
   </div> <!-- 功能表1 end -->
</div> 
<div class="w3-pale-red w3-center" >
   <div class="w3-bar w3-blue" ><!-- 功能表2 start -->
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-wk_Calendar_r0.asp?worker=<%=worker%>')">回日曆表</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-work_query.asp?worker=<%=worker%>')">工作查詢</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-work_add.asp?worker=<%=worker%>')">工作新增</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-wk_pj_list.asp')" title="專案工作">專案工作</button>
      <button class="w3-bar-item w3-button w3-mobile" style="" onclick="url_show('zec-2_admin_main.asp')" title="後台管理">後台管理</button>
   </div><!-- 功能表2 end -->
   <div class="w3-row w3-center " ><!-- 內容 start -->
      <div class="w3-responsive">
      <table class="w3-table-all" border=1>
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <col style="width:14.2857%;">
         <tr>
           <td colspan=5 style="text-align:center;">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【<<】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【<】</button>
            <%=p_year%>年<%=p_month%>月
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【今日=<%=pn_date%>】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【>】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【>>】</button>
           </td>
           <td colspan=2 style="text-align:center;">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【月】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【週】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【日】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;">【表】</button>
           </td>
         </tr>
         <tr>
           <th style="text-align:center;">星期日</th>
           <th style="text-align:center;">星期一</th>
           <th style="text-align:center;">星期二</th>
           <th style="text-align:center;">星期三</th>
           <th style="text-align:center;">星期四</th>
           <th style="text-align:center;">星期五</th>
           <th style="text-align:center;">星期六</th>
         </tr>
         <tr>
           <td>Jill</td>
           <td>Smith</td>
           <td>50</td>
           <td>50</td>
           <td>50</td>
           <td>50</td>
           <td>50</td>
         </tr>
         <tr>
           <td>Eve</td>
           <td>Jackson</td>
           <td>94</td>
           <td>94</td>
           <td>94</td>
           <td>94</td>
           <td>94</td>
         </tr>
         <tr>
           <td>Adam</td>
           <td>Johnson</td>
           <td>67</td>
           <td>67</td>
           <td>67</td>
           <td>67</td>
           <td>67</td>
         </tr>
      </table>
      </div>
   </div><!-- 內容 end -->
</div>
<!--內容-->
<div class="w3-red w3-center" >
   <div class="w3-row w3-center " >

<!--      <iframe id="ifrm_content" name="ifrm_content" src="zec-wk_Calendar_r0.asp?worker=<%=worker%>" style="border:2px;width:100%;height:100%;"></iframe>	-->
   </div>
</div>
<!--頁尾-->
<div class="footer w3-brown w3-center"  style="height:40px;overflow: hidden;margin:0px;padding:0px;">
  <h6>　　　版權所有　　　<strong>喬大地產開發股份有限公司</strong>　　　2021　　　</h6>
</div>

<script language="JavaScript">
    function url_new(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        //window.location.href = pp_url; //原頁面更新
        window.open(pp_url) ; //開啟新頁面
        return true;
    }   
    function url_show(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        window.location.href = pp_url; //原頁面更新
        //window.open(pp_url) ; //開啟新頁面
        return true;
    }   
    function content_show(pp_url){
        var iframe1=document.getElementById("ifrm_content");
        iframe1.src=pp_url;
        return true;
    }    

</script>

</body>
</html>