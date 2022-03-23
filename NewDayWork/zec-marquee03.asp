<%@ Language=VBScript CODEPAGE=950 %>
<!DOCTYPE html>
<html>

<!--head區段設-->
<head>
<style>
/*在head區段內先設定一個marquee樣式，用class和id都沒差吧*/
.marquee {
/*行高設定*/
 height: 40px;
/*隱藏多出來的文字*/
 overflow: hidden; 
/*隱藏多出來的行*/
 position: relative;
}

/*文字外觀與動畫執行的設定*/
.marquee ul {
/*清除ul的項目點點*/
 list-style-type: none;
/*動做設定：動畫名稱、要跑多久、運動模式、次數*/
 animation-name: maruqee;
 animation-duration:15s;
 animation-timing-function: linear;
/*執行次數：infinite（無限重複）、3(指定3次)*/
 animation-iteration-count:infinite;
/*給這個屬性才會有區域，不然就只有內容總長*/    
 position: absolute;
}

/*動畫行為的安排*/
@keyframes maruqee {
/*動作的起始位置*/
 from {
  left: 100%;
 }
/*動作的結束位置*/
 to {
  left: 0%;
 }
}
</style>
</head>

<body>

<!--指定div的id為marqee-->
  <div class="marquee">

<!--把文字放入ul/li列表項目語法中，使用程式就自己換純文字區域-->
     <ul>
         <li>2020.2.4(一)：公告事項呈現區域，無公告則隱藏
         至此已經搞得我精疲力盡了，用safari開啟，果然有在跑，這次總算是happy ending的大結局了。
         至此已經搞得我精疲力盡了，用safari開啟，果然有在跑，這次總算是happy ending的大結局了。</li>
     </ul>
  </div>
</body>
</html>
