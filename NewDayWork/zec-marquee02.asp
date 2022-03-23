<%@ Language=VBScript CODEPAGE=950 %>
<!DOCTYPE html>
<html lang="zh-Hant-TW">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
.marquee {
    position: relative;
    overflow: hidden;
    --offset: 20vw;
    --move-initial: calc(-25% + var(--offset));
    --move-final: calc(-50% + var(--offset));
}

.marquee__inner {
    width: fit-content;
    display: flex;
    position: relative;
    transform: translate3d(var(--move-initial), 0, 0);
    animation: marquee 5s linear infinite;
    /*animation-play-state: paused;*/
    animation-play-state: running;
}

.marquee span {
    font-size: 10vw;
    padding: 0 2vw;
}

.marquee:hover .marquee__inner {
    /*animation-play-state: running;*/
    animation-play-state: paused;
}

@keyframes marquee {
    0% {
        transform: translate3d(var(--move-initial), 0, 0);
    }

    100% {
        transform: translate3d(var(--move-final), 0, 0);
    }
}
</style>
</head>
<body>
<div class="marquee">
	<div class="marquee__inner" aria-hidden="true">
		<span>Showreel1</span>
		<span>test</span>
		<span>gogogo</span>
		<span>zzzzz</span>
	</div>
</div>
</body>
</html>