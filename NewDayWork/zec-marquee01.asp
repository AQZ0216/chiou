<%@ Language=VBScript CODEPAGE=950 %>
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
.marquee {
    width: 100%;
    margin: 0 auto;
    white-space: nowrap;
    overflow: hidden;
    box-sizing: border-box;
    display: inline-flex;    
}

.marquee span {
    display: flex;        
    flex-basis: 100%;
    animation: marquee-reset;
    animation-play-state: paused;                
}

.marquee:hover> span {
    animation: marquee 2s linear infinite;
    animation-play-state: running;
}

@keyframes marquee {
    0% {
        transform: translate(0%, 0);
    }    
    50% {
        transform: translate(-100%, 0);
    }
    50.001% {
        transform: translate(100%, 0);
    }
    100% {
        transform: translate(0%, 0);
    }
}
@keyframes marquee-reset {
    0% {
        transform: translate(0%, 0);
    }  
}
</style>
</head>
<body>
<span class="marquee">
   <span>
       When I had journeyed half of our life's way, I found myself
       within a shadowed forest, for I had lost the path that 
       does not stray. – (Dante Alighieri, <i>Divine Comedy</i>. 
       1265-1321)
   </span>
</span>
</body>
</html>