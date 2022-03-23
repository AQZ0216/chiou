<% @codepage=950%>
<!-- #Include file = "./include/array_worker_crew.inc" -->
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
	'讀取人員姓名
	worker = Request("worker")
	wk_id=Request("wk_id")
   wk_chk=Request("wk_chk")
%>
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
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id
rstObj1.open strSQL_show,conDB,3,1
'讀取資料
undo_date1=rstObj1.fields("undo_date1")
doing_date1=rstObj1.fields("doing_date1")
done_date1=rstObj1.fields("done_date1")
p_wkid=rstObj1.fields("wk_id")
wk_item=rstObj1.fields("wk_item")
wk_content=rstObj1.fields("wk_content")
wk_order=rstObj1.fields("wk_order")
wk_doer=rstObj1.fields("wk_doer")
wk_checker=rstObj1.fields("wk_checker")
wk_undoer=rstObj1.fields("wk_undoer")
wk_class=rstObj1.fields("wk_class")
wk_group=rstObj1.fields("wk_group")
wk_exe=rstObj1.fields("wk_exe")           '執行人員
wk_att=rstObj1.fields("wk_att")           '出席人員
wk_pjn=rstObj1.fields("pj_02")   '專案名稱
pwk_password=rstObj1.fields("wk_password")   '加密文字
wk_headline=rstObj1.fields("headline")'跑馬燈

%>
<%
'關閉資料集
rstObj1.Close
'重設資料變數 
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數 
set conDB=Nothing 
%>
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<title>【<%=worker%>】工作管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="./css/w3-cht.css">
<link rel="stylesheet" href="./css/font-awesome.min.css">

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
   
      <table class="w3-table-all" ><!-- 功能表 start -->
         <col style="width:100%;background-color:#ffdddd;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <td style="text-align:center;height:55px;">
<% 
if isnull(wk_headline) then wk_headline=0
if cint(wk_headline) > 5 then 
%>
<img src="./img/gnome_chess.png" width=32 onclick="parent.content.location.href='0_wk_headline_off_20140728.asp?wk_id=<%=wk_id%>'" title="已在跑馬燈訊息中，移出跑馬燈訊息">
<% else %>
<img src="./img/gnome_chess_d.png" width=32 onclick="parent.content.location.href='0_wk_headline_on_20140728.asp?wk_id=<%=wk_id%>'" title="不在跑馬燈訊息中，轉入跑馬燈訊息">
<% end if %>           
            <button class="w3-button w3-white w3-xlarge " style="padding:2px;margin:0px;" >單一工作項目顯示</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="回上一頁" onclick="javascript:history.go(-1)">回上一頁</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="修改工作" onclick="url_show_confirm('zec-work_edit_r1.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">修改工作</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="刪除工作" onclick="url_show_confirm('zec-work_del_r1.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">刪除工作</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="完成工作" onclick="url_show_confirm('zec-work_del_r1.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">完成工作</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="重新公告" onclick="url_show_confirm('zec-work_readd.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">重新公告</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="列印內容" onclick="url_open('zec-work_print_si.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">列印內容</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="轉為專案" onclick="url_show_confirm('zec-work_gpchg_special.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">轉為專案</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="轉為一般" onclick="url_show_confirm('zec-work_gpchg_normal.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">轉為一般</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="上傳附件" onclick="url_show_confirm('zec-1_ulf_form.asp?worker=<%=worker%>&wk_id=<%=p_wkid%>')">上傳附件</button>
           </td>
         </tr>
      </table><!-- 功能表 end -->
      
<div class="w3-container w3-center" style="margin:0px;padding:0px;height:520px;overflow:auto;"><!-- 工作項目表 start -->
      <div class="w3-responsive" ><!-- div w3-responsive start -->
      <table class="w3-table-all" >
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <col style="width:14.2857%;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <th style="text-align:center;background-color:#ffdddd;border:1px solid #000;">工作群組</th>
           <th style="text-align:center;background-color:#ddffdd;border:1px solid #000;">專案名稱</th>
           <th style="text-align:center;background-color:#ffdddd;border:1px solid #000;">工作編號</th>
           <th style="text-align:center;background-color:#ddffdd;border:1px solid #000;">工作分類</th>
           <th style="text-align:center;background-color:#ffdddd;border:1px solid #000;">公告者</th>
           <th style="text-align:center;background-color:#ddffdd;border:1px solid #000;">公告日期</th>
           <th style="text-align:center;background-color:#ffdddd;border:1px solid #000;">執行日期</th>
         </tr> 
         <tr style="border:1px solid #000;" >
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;"><%=wk_group%></td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;"><%=wk_pjn%></td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;"><%=wk_id%></td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;"><%=wk_class%></td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;"><%=wk_order%></td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;"><%=undo_date1%></td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;"><%=doing_date1%></td>
         </tr>
         <tr style="border:1px solid #000;" >
            <td>執行人員</td>
            <td colspan="6"><%=wk_exe%></td>
         </tr>
         <tr style="border:1px solid #000;" >
            <td>出席人員</td>
            <td colspan="6"><%=wk_att%></td>
         </tr>
         <tr style="border:1px solid #000;" >
            <td>知會人員</td>
            <td colspan="6"><%=wk_doer%></td>
         </tr>
         <tr style="border:1px solid #000;" >
         	<td>完成人員</td>
         	<td colspan="6"><%=wk_checker%></td>
         </tr>
         <tr style="border:1px solid #000;" >
         	<td>未完成人員</td>
         	<td colspan="6"><%=wk_undoer%></td>
         </tr>
         <tr style="border:1px solid #000;" >
         	<td align="right">主旨：</td>
         	<td colspan="6"><%=wk_item%></td>
         </tr>
         <tr style="border:1px solid #000;" >
         	<td align="right" valign="top">執行內容：</td>
         	<td colspan="6" style="padding:0px;margin:0px;">
               <% 
                  pp_wk_content=replace(wk_content,chr(13),"<br>",1,-1,1)
                  'response.write pp_wk_content
               %>
               <div style="margin:0px;padding:0px;height:200px;overflow:auto;">
               <%=pp_wk_content%>
               </div>
          	</td>
         </tr>
         <tr style="border:1px solid #000;" data-ng-bind="">
         	<td align="right"><font color="red">加密文字：</font></td>
         	<td colspan="6"><%=pwk_password%></td>
         </tr>
      </table>     
      </div><!-- div w3-responsive end -->
      <div class="w3-responsive" ><!-- div w3-responsive start -->
<%
'附加檔案列表
' 連結Access資料庫daywork.mdb
DBpath=Server.MapPath("./database/attach_file.mdb")
strCon="Driver={Microsoft Access Driver (*.mdb)};DBQ="&DBpath
'建立資料庫連結物件
set conDB= Server.CreateObject("ADODB.Connection")
'連結資料庫	
conDB.Open strCon
'開啟資料表名稱
tb_name="file_data"
'建立資料庫存取物件	
set rstObj1=Server.CreateObject("ADODB.Recordset")
strSQL_show="Select * from " & tb_name & " where wk_id =" & wk_id &" and del_ok = false order by fl_date desc"
rstObj1.open strSQL_show,conDB,3,1
totalput=rstObj1.recordcount
if totalput=0 then
else
%>
<table class="w3-table-all" >
<col style="width:80px;border:1px solid #000;">
<col style="border:1px solid #000;">
<col style="width:300px;border:1px solid #000;">
<col style="width:90px;border:1px solid #000;">
<col style="width:150px;border:1px solid #000;">
<tr style="border:1px solid #000;" >
<td colspan=5>附件列表</td>
</tr>
<tr style="background-color:#ffdddd;border:1px solid #000;">
<th>序號</th>
<th>檔案說明</th>
<th>檔案名稱  [上傳者]</th>
<th>建檔日期</th>
<th>功能</th>
</tr>
<%
	'列出資料項目
	rstobj1.MoveFirst
	for fi=1 to totalput
	'讀取資料
		pfl_id=rstObj1.fields("fl_id")
		pwk_id=rstObj1.fields("wk_id")
		pfl_name=rstObj1.fields("fl_name")
		pfl_item=rstObj1.fields("fl_item")
		pfl_inputer=rstObj1.fields("fl_inputer")
		pfl_history= rstObj1.fields("fl_history")
		pfl_date=rstObj1.fields("fl_date")
		str_none=pwk_id&"_"
		str_pfl_name=right(pfl_name,len(pfl_name)-len(pwk_id)-1)
%>
<tr style="border:1px solid #000;">
<td style="text-align:center;"><%=fi%></td>
<td >
<a href="./zec-1_ulf_item_edit.asp?worker=<%=worker%>&wk_id=<%=pwk_id%>&fl_id=<%=pfl_id%>" target="_self" title="修改檔案說明。" ><img src="./img/change.png" style="vertical-align:middle;height:16px;cursor:hand;border:0;" ></a>
<%=pfl_item%>
</td>
<td ><a href="./file_att/<%=pfl_name%>" target="_blank" title="<%=pfl_history%>"><%=str_pfl_name%></a>  [<%=pfl_inputer%>]</td>
<td ><%=pfl_date%></td>
<td >
<button class="w3-button w3-blue" style="padding:2px;margin:0px;" onclick="url_show_confirm('zec-1_ulf_form.asp?worker=<%=worker%>&wk_id=<%=pwk_id%>')" title="工作項目 [ wk_id=<%=pwk_id%> ] 新增檔案。">【新】</button>
<button class="w3-button w3-blue" style="padding:2px;margin:0px;" onclick="url_show_confirm('zec-1_ulf_del.asp?worker=<%=worker%>&wk_id=<%=pwk_id%>&fl_id=<%=pfl_id%>')" title="將檔案刪除。">【刪】</button>
</td>
</tr>
<%
	'移到下一筆記錄
		rstObj1.MoveNext
		if rstObj1.EOF=True then exit for
	next	
%>

</table>
<%
end if
'關閉資料集
rstObj1.Close
'重設資料變數
set rstObj1=Nothing
'關閉資料庫 
conDB.Close
'重設物件變數
set conDB=Nothing
%>
      </div><!-- div w3-responsive end -->
      
</div><!-- 工作項目表 end -->

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
    function url_show_confirm(pp_url){
        //var iframe1=document.getElementById("ifrm_milestone");
        //iframe1.src=pp_url;
        //
        aws=confirm("確定執行【"+pp_url+"】嗎？？");
        if (aws)
        {
            window.location.href = pp_url; //原頁面更新
            //window.open(pp_url) ; //開啟新頁面
            return true;         
        }
        //else
        //{
        // alert("取消執行【"+pp_url+"】！！");
        //}       
    }       
    function content_show(pp_url){
        var iframe1=document.getElementById("ifrm_content");
        iframe1.src=pp_url;
        return true;
    }    

</script>

</body>
</html>