<% @codepage=950%>
<!-- #Include file = "./include/array_worker_crew.inc" -->
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
		       pstr_hdman = pstr_hdman & "<span style='font-size:14px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man & "</span><br>"
   	           end if
		else
		  pstr_hdman = pstr_hdman & "<span style='font-size:14px;font-weight:bold;color:"& f_color &";'>" & hd_img & hd_man &"("& hd_hrs&")&nbsp;</span><br>"
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
<%
'查詢是否有附件
Function exist_attach(pwk_id)
      ' 連結Access資料庫daywork.mdb
      DBpath_fe=Server.MapPath("./database/attach_file.mdb")
      strCon_fe="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath_fe
      '建立資料庫連結物件
      set conDB_fe= Server.CreateObject("ADODB.Connection")
      '連結資料庫	
      conDB_fe.Open strCon_fe
      '開啟資料表名稱
      tb_name_fe="file_data"
      '建立資料庫存取物件	
      set rstObj1_fe=Server.CreateObject("ADODB.Recordset")
      strSQL_show_fe="Select * from " & tb_name_fe & " where del_ok = false and wk_id = "& pwk_id &" order by wk_id desc"
      rstObj1_fe.open strSQL_show_fe,conDB_fe,3,1
      totalput_fe=rstObj1_fe.recordcount
      '關閉資料集
      rstObj1_fe.Close
      '重設資料變數
      set rstObj1_fe=Nothing
      '關閉資料庫 
      conDB_fe.Close
      '重設物件變數
      set conDB_fe=Nothing
      exist_attach=totalput_fe
End Function

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
<!-- 標頭開始 -->
<!-- #Include file = "./include/zec-header_r1.inc" -->
<!-- 標頭結束 -->
<!-- 內文開始 -->
<div class="vt-container w3-pale-red w3-center" >
   <div class="w3-row w3-center " ><!-- 內容 start -->
<%
datecode=request("datecode")
p_year=year(datecode)'年
p_month=month(datecode)'月
p_day=day(datecode)'日
pn_date=dateserial(p_year,p_month,p_day)'查詢日期
pn_weekday=Weekday(pn_date)'星期幾

      Select Case pn_weekday
         Case 1    
            str_wk="星期日"
            div_background_c="#ffdddd"    'w3-pale-red
            div_border_c="#000"        'w3-pale-red
         Case 2    
            str_wk="星期一"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
         Case 3    
            str_wk="星期二"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
        Case 4    
            str_wk="星期三"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
         Case 5    
            str_wk="星期四"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
         Case 6    
            str_wk="星期五"
            div_background_c="#ddffdd"    'w3-pale-green
            div_border_c="#000"        'w3-pale-green
        Case 7    
            str_wk="星期六"
            div_background_c="#ffdddd"    'w3-pale-red
            div_border_c="#000"        'w3-pale-red
         Case Else     
      End Select 

%>
      <table class="w3-table-all" >
         <col style="width:100%;background-color:#ffdddd;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <td style="text-align:center;height:55px;">
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="上一年" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=cint(p_year)-1%>&p_month=<%=p_month%>')" >【<<】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="上一月" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=p_year%>&p_month=<%=cint(p_month)-1%>')" >【<】</button>
            <button class="w3-button w3-white w3-xlarge " style="padding:2px;margin:0px;" >【<%=p_year%>年<%=p_month%>月<%=p_day%>日】單日工作項目</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="下一月" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=p_year%>&p_month=<%=cint(p_month)+1%>')">【>】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="下一年" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=cint(p_year)+1%>&p_month=<%=p_month%>')">【>>】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="回<%=year(date())%>年<%=month(date())%>月" onclick="url_show('zec-work_month_r1.asp?worker=<%=worker%>&p_year=<%=year(date())%>&p_month=<%=month(date())%>')">【今月：<%=year(date())%>年<%=month(date())%>月】</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="回<%=year(date())%>年<%=month(date())%>月<%=day(date())%>日" onclick="url_show('zec-work_day_r1.asp?worker=<%=worker%>&datecode=<%=date()%>')">【今日：<%=date()%>】</button>
           </td>
         </tr>
      </table>
<div class="w3-container" style="margin:0px;padding:0px;height:520px;overflow:auto;"><!-- 日曆表 start -->
      <div class="w3-responsive">
      <table class="w3-table-all" >
         <col style="width:100%;background-color:<%=div_background_c%>;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <th style="text-align:center;background-color:<%=div_background_c%>;border:1px solid #000;">
           <span class="w3-button" style="font-size:24px;padding:0px;margin:0px;">【<%=p_year%>年<%=p_month%>月<%=p_day%>日】<%=str_wk%></span>
           <span class="w3-button w3-red" title="新增工作" style="font-size:24px;padding:0px;margin:0px;" onclick="url_show('zec-work_add_r1.asp?datecode=<%=pn_date%>&worker=<%=worker%>')" > 【新增】 </span>
           </th>
         </tr>
<!-- 開啟資料庫 -->
<%
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/daywork.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="work_data"
%>   
<%
      pndate=pn_date '顯示的工作日期
      str_hdman=hd_man(pndate)'休假字串
'-------------------------str_allwork----------------------------
      str_allwork="" '工作字串
      '建立資料庫存取物件	
      set rstObj1=Server.CreateObject("ADODB.Recordset")
      strSQL_show="Select * from " & tb_name & " where doing_date1 = #"& pndate &"# and wk_undoer like '%"& worker &"%' order by wk_item asc , wk_id asc"
      rstObj1.open strSQL_show,conDB,3,1
      tot_d1=rstObj1.recordcount
      if not rstObj1.EOF then
         rstObj1.Movefirst
         for di = 1 to tot_d1
            p_wkid=rstObj1.fields("wk_id")
            p_wkitem=rstObj1.fields("wk_item") 
            '----------------------------------------
               wk_headline=rstObj1.fields("headline")  '跑馬燈
               '檢查是否有附件 exist_attach(wk_id)
               attach_no=exist_attach(p_wkid)
               if attach_no=0 then
                  str_colors="color:#000000;"
               else
                  str_colors="color:#0000FF;"
               end if
               if rstObj1.fields("wk_password")="" or isnull(rstObj1.fields("wk_password")) then
               else
                  str_colors="color:#0000FF;"
               end if
					p_nexe=rstObj1.fields("wk_exe")	'執行人員
					if Instr(1, p_nexe, worker, 1)>0 or Instr(1, p_nexe, "全體", 1)>0 then
						str_bgc="background-color:#99FF99;"	
					else
						str_bgc=""
					end if            
            '----------------------------------------
           
            str_allwork = str_allwork & "<span style='font-size:14px;"& str_bgc & str_colors &"' >" & di &"、<a href='zec-work_show_r1.asp?wk_id="& p_wkid &"&worker="& worker&"' style='text-decoration: none;"& str_colors &"' >" & p_wkitem &"</a></span><br>" 
            rstObj1.MoveNext
            if rstObj1.EOF=true then exit for
         next
      else
      end if
      '關閉資料集
      rstObj1.Close
      '重設資料變數 
      set rstObj1=Nothing
'-------------------------str_allwork----------------------------
%>          
   <tr>
      <td style="background-color:<%=div_background_c%>;border-color:<%=div_border_c%>;" >
         <%=str_allwork%><%=str_hdman%>       
      </td>
   </tr>
   </table>
      </div>
<%

'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%> 	      
</div><!-- 日曆表 end -->

   </div><!-- 內容 end -->
</div>
<!-- 內文結束 -->
<!-- 頁尾開始 -->
<!-- #Include file = "./include/zec-footer_r1.inc" -->
<!-- 頁尾結束 -->

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