<% @codepage=950%>
<!-- #Include file = "./include/array_worker_crew.inc" -->
<!-- #Include file = "./include/array_wkclass.inc" -->
<!-- #Include file = "./include/array_wkgroup.inc" -->
<!-- #Include file = "./include/workinput.inc" -->
<!-- #Include file = "./misc_data/array_place.inc" -->	
<!-- #Include file = "./misc_data/array_thing.inc" -->
<!-- #Include file = "./include/array_pjn.inc" -->
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
'Generates date in yyyy-mm-dd format
Function GetFormattedDate(setDate)
   strDate = CDate(setDate)
   strDay = DatePart("d", strDate)
   strMonth = DatePart("m", strDate)
   strYear = DatePart("yyyy", strDate)
   If strDay < 10 Then
     strDay = "0" & strDay
   End If
   If strMonth < 10 Then
     strMonth = "0" & strMonth
   End If
   GetFormattedDate = strYear & "-" & strMonth & "-" & strDay
End Function
%>
<%
	'讀取資料
	worker = request("worker") '讀取人員
	datecode=request("datecode")'讀取日期
%>
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<title>【<%=worker%>】工作--新增</title>
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
<!-- ----------------------------------------------------------------- 內容start -----------------------------------------------------------------  -->   
<form name="form1" action="zec-work_add_ok_r1.asp" method="post" >
      <table class="w3-table-all" ><!-- 功能表 start -->
         <col style="width:100%;background-color:#ffdddd;border:1px solid #000;">
         <tr style="border:1px solid #000;">
           <td style="text-align:center;height:55px;">
            <button class="w3-button w3-white w3-xlarge " style="padding:2px;margin:0px;" >【<%=worker%>】單一工作項目新增</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="回上一頁" onclick="javascript:history.go(-2)">回上一頁</button>
            <button class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="確定新增" onclick="">確定新增</button>
            <input type="reset" class="w3-button w3-blue w3-medium " style="padding:4px;margin:2px;" title="重設資料" value="重設資料" >
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
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;padding:0px;margin:0px;"><!-- 工作群組 -->
            	<select name="wk_group" style="width:100%;padding:0px;margin:0px;height:100%;">
            <%
            		response.write "<option value='"&wk_group_a(0)&"' selected>"&wk_group_a(0)
            	for i=2 to wk_group_no
            		response.write "<option value='"&wk_group_a(i-1)&"'>"&wk_group_a(i-1)
            	next
            %>
            	</select>
            </td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;padding:0px;margin:0px;"><!-- 專案名稱 -->
            	<select name="wk_pjn" style="width:100%;padding:0px;margin:0px;height:100%;" >
            <%
            		response.write "<option value='0' selected>"
            		'response.write "<option value='"&pjnid_a(0)&"' >"&pjn_a(0)
            	for i=1 to pjn_no
            		response.write "<option value='"&pjnid_a(i-1)&"，"&pjn_a(i-1)&"'>"&pjn_a(i-1)
            	next
            %>
            	</select>
            </td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;padding:0px;margin:0px;"><!-- 工作編號 -->
               自動編號
            </td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;padding:0px;margin:0px;"><!-- 工作分類 -->
            	<select name="wk_class" style="width:100%;padding:0px;margin:0px;height:100%;" >
            <%
            	for i=1 to wk_class_no
            		response.write "<option value='"&wk_class_a(i-1)&"'>"&wk_class_a(i-1)
            		 if wk_class_a(i-1)="Z" then response.write "-不要完成"
            	next
            %>
            	</select>
            </td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;padding:0px;margin:0px;"><!-- 公告者 -->
               <input type='text' name='wk_order' value='<%=worker%>' style="width:100%;" readonly>
            </td>
            <td style="text-align:center;background-color:#ddffdd;border:1px solid #000;padding:0px;margin:0px;"><!-- 公告日期 -->
               <input type="date" id="datePicker" name='undo_date1' value="<%=GetFormattedDate(datecode)%>" style="width:100%;" >
            </td>
            <td style="text-align:center;background-color:#ffdddd;border:1px solid #000;padding:0px;margin:0px;"><!-- 執行日期 -->
               <input type="date" name="doing_date1" value="<%=GetFormattedDate(date())%>" style="width:100%;">
            </td>
         </tr>
         <tr style="border:1px solid #000;" >
            <td>執行人員</td>
            <td colspan="6" style="text-align:left;background-color:#ffdddd;border:1px solid #000;">
            	<input type='text' name='wk_exe' value='' style="width:50%;" readonly title="執行人員請採用右方下拉選單輸入！！！" onkeydown="javascript:if(window.event.keyCode==8) return false;">
		<SELECT name="exemen_w" onchange="exeadd()">
		<option value="" selected>請選擇人員</option>
		<option value="clear" >清除人員</option>
			<option value="全體人員" >全體人員</option>
		<option value="業務全體" >業務全體</option>
		<option value="內勤全體" >內勤全體</option>
	<%
		for i=1 to worker_no
			response.write "<option value='" & worker_a(i-1) & "'>" & worker_a(i-1) &"</option>"
		next
	%>
		</SELECT>

		<SELECT name="exemen_dp" onchange="exeadddp()">
			<option value="" selected>部門選擇</option>
			<option value="clear" >清除人員</option>
			<option value="<%=stra_dp01%>" >總經理室</option>
			<option value="<%=stra_dp02%>" >管理部</option>
			<option value="<%=stra_dp03%>" >企劃部</option>
			<option value="<%=stra_dp04%>" >業務部</option>
			<option value="<%=stra_dp05%>" >法務部</option>
			<option value="<%=stra_dp06%>" >財務部</option>
			<option value="<%=stra_dp07%>" >資訊部</option>
			<option value="<%=stra_dp08%>" >建設部</option>
			<option value="<%=stra_dp10%>" >我家農業</option>
			<option value="<%=stra_dpa1%>" >業一</option>
			<option value="<%=stra_dpa2%>" >業二</option>
			<option value="<%=stra_dpa3%>" >業Three</option>
			<option value="<%=stra_dpa4%>" >YES</option>
			<option value="<%=stra_dpa5%>" >冒泡業八</option>
		</SELECT>	
				(請輸入執行參與人員)	
            </td>
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
     
</div><!-- 工作項目表 end -->
</form>
<!-- -----------------------------------------------------------------內容 end----------------------------------------------------------------- -->   
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