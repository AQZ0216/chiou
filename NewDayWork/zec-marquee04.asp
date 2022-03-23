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
            text-overflow: ellipsis; /*�W�X������...�N��*/
            white-space:nowrap;/*�j���r�b�@�椺���*/
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
        <span> 0102�ܦ��w�g�d�o�ں�h�O�ɤF�A��safari�}�ҡA�G�M���b�]�A�o���`��Ohappy ending���j�����F�C
        0202�ܦ��w�g�d�o�ں�h�O�ɤF�A��safari�}�ҡA�G�M���b�]�A�o���`��Ohappy ending���j�����F�C
        0302�ܦ��w�g�d�o�ں�h�O�ɤF�A��safari�}�ҡA�G�M���b�]�A�o���`��Ohappy ending���j�����F�C</span>
    </div>
</body>

</html>