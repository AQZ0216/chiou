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
  overflow: hidden;
  box-sizing: border-box;
}

.marquee span {
  display: inline-block;
  width: max-content;

  padding-left: 100%;
  /* show the marquee just outside the paragraph */
  will-change: transform;
  animation: marquee 15s linear infinite;
}

.marquee span:hover {
  animation-play-state: paused
}


@keyframes marquee {
  0% { transform: translate(0, 0); }
  100% { transform: translate(-100%, 0); }
}


/* Respect user preferences about animations */

@media (prefers-reduced-motion: reduce) {
  .marquee span {
    animation-iteration-count: 1;
    animation-duration: 0.01; 
    /* instead of animation: none, so an animationend event is 
     * still available, if previously attached.
     */
    width: auto;
    padding-left: 0;
  }
}
</style>
</head>
<body>
<p class="marquee">
   <span>
       When I had journeyed half of our life's way, I found myself
       within a shadowed forest, for I had lost the path that 
       does not stray. – (Dante Alighieri, <i>Divine Comedy</i>. 
       1265-1321)
   </span>
   </p>
</body>
</html>