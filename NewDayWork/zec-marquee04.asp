<%@ Language=VBScript CODEPAGE=950 %>
<!DOCTYPE html>
<html>

<body>
    <style style="text/css">
        .marquee {
            height: 50px;
            overflow: hidden;
            position: relative;
            background: #fefefe;
            color: #333;
            border: 1px solid #4a4a4a;
        }
        
        .marquee span {
            text-overflow: ellipsis; /*超出部分用...代替*/
            white-space:nowrap;/*強制文字在一行內顯示*/
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
    </style>
<body>
    <div class="marquee">
        <span> 0102至此已經搞得我精疲力盡了，用safari開啟，果然有在跑，這次總算是happy ending的大結局了。
        0202至此已經搞得我精疲力盡了，用safari開啟，果然有在跑，這次總算是happy ending的大結局了。
        0302至此已經搞得我精疲力盡了，用safari開啟，果然有在跑，這次總算是happy ending的大結局了。</span>
    </div>
</body>

</html>