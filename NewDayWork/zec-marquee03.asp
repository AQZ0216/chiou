<%@ Language=VBScript CODEPAGE=950 %>
<!DOCTYPE html>
<html>

<!--head�Ϭq�]-->
<head>
<style>
/*�bhead�Ϭq�����]�w�@��marquee�˦��A��class�Mid���S�t�a*/
.marquee {
/*�氪�]�w*/
 height: 40px;
/*���æh�X�Ӫ���r*/
 overflow: hidden; 
/*���æh�X�Ӫ���*/
 position: relative;
}

/*��r�~�[�P�ʵe���檺�]�w*/
.marquee ul {
/*�M��ul�������I�I*/
 list-style-type: none;
/*�ʰ��]�w�G�ʵe�W�١B�n�]�h�[�B�B�ʼҦ��B����*/
 animation-name: maruqee;
 animation-duration:15s;
 animation-timing-function: linear;
/*���榸�ơGinfinite�]�L�����ơ^�B3(���w3��)*/
 animation-iteration-count:infinite;
/*���o���ݩʤ~�|���ϰ�A���M�N�u�����e�`��*/    
 position: absolute;
}

/*�ʵe�欰���w��*/
@keyframes maruqee {
/*�ʧ@���_�l��m*/
 from {
  left: 100%;
 }
/*�ʧ@��������m*/
 to {
  left: 0%;
 }
}
</style>
</head>

<body>

<!--���wdiv��id��marqee-->
  <div class="marquee">

<!--���r��Jul/li�C���ػy�k���A�ϥε{���N�ۤv���¤�r�ϰ�-->
     <ul>
         <li>2020.2.4(�@)�G���i�ƶ��e�{�ϰ�A�L���i�h����
         �ܦ��w�g�d�o�ں�h�O�ɤF�A��safari�}�ҡA�G�M���b�]�A�o���`��Ohappy ending���j�����F�C
         �ܦ��w�g�d�o�ں�h�O�ɤF�A��safari�}�ҡA�G�M���b�]�A�o���`��Ohappy ending���j�����F�C</li>
     </ul>
  </div>
</body>
</html>
