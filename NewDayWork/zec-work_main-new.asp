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
   <div class="w3-row w3-pale-green ">
      <!--功能表1-->
      <button class="w3-button w3-red w3-medium w3-round" onclick="url_new('http://192.168.0.11/chiou/att2000/5_card_query.asp')" title="出勤時間" style="padding:4px;margin:2px;">出勤時間</button>
      <button class="w3-button w3-red w3-medium w3-round" onclick="url_new('../holiday/hd_ps_year_list.asp?wkr_id=<%=pwkr_id%>')" title="休假資料" style="padding:4px;margin:2px;">休假資料</button>
      <button class="w3-button w3-red w3-medium w3-round" onclick="url_new('http://1.34.48.166:90/firstpage.asp')" title="球場日曆" style="padding:4px;margin:2px;">球場日曆</button>
      <button class="w3-button w3-red w3-medium w3-round" onclick="url_new('http://60.251.159.62:6980/build/daywork/firstpage.asp?paswd=28283939')" title="建設部" style="padding:4px;margin:2px;">建設部</button>
   </div>  
</div>
<!--內容-->
<div class="w3-brown w3-center" >
   <div class="w3-row w3-pale-blue ">
      <!--功能表2-->
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-wk_Calendar_all.asp?worker=<%=worker%>')" title="回日曆表" style="padding:4px;margin:2px;">回日曆表</button>
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-work_query.asp?worker=<%=worker%>')" title="工作查詢" style="padding:4px;margin:2px;">工作查詢</button>
      <button class="w3-button w3-blue w3-medium w3-round" onclick="content_show('zec-work_add.asp?worker=<%=worker%>')" title="工作新增" style="padding:4px;margin:2px;">工作新增</button>
   </div>     
   <div class="w3-row w3-center w3-pale-blue" style="width:100%;height:550px;overflow: hidden;">
      <iframe id="ifrm_content" name="ifrm_content" src="zec-wk_Calendar_new.asp?worker=<%=worker%>" style="border:0px;width:100%;height:100%;"></iframe>	
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