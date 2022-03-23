
 <head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>Karypov PP</title>
  <link rel="stylesheet" type="text/css" href="./css/style.css">
  <center><h3><p>Форма</p></h3>
  <style>
	/* cyrillic */
	@font-face {
	font-family: "CormorantGaramond-Regular";
	src: url("./CormorantGaramond-Regular.ttf");
	}
	p {
	font-family: "CormorantGaramond-Regular";
	}
      #ready {
        display: none;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
      }
      #okno {
        width: 300px;
        height: 50px;
        text-align: center;
        padding: 15px;
        border: 3px solid #008000;
        border-radius: 10px;
        color: #008000;
        position: absolute;
        top: 0;
        right: 0;
        bottom: 0;
        left: 0;
        margin: auto;
      }
      #ready:target {display: block;}
    </style>
 </head>
 <body>
	<a href="#" id="ready">
      <div id="okno">
        Готово!
      </div>
    </a>
<!-- Yandex.Metrika counter -->
<script type="text/javascript" >
   (function(m,e,t,r,i,k,a){m[i]=m[i]||function(){(m[i].a=m[i].a||[]).push(arguments)};
   m[i].l=1*new Date();k=e.createElement(t),a=e.getElementsByTagName(t)[0],k.async=1,k.src=r,a.parentNode.insertBefore(k,a)})
   (window, document, "script", "https://mc.yandex.ru/metrika/tag.js", "ym");

   ym(87901924, "init", {
        clickmap:true,
        trackLinks:true,
        accurateTrackBounce:true
   });
</script>
<noscript><div><img src="https://mc.yandex.ru/watch/87901924" style="position:absolute; left:-9999px;" alt="" /></div></noscript>
<!-- /Yandex.Metrika counter -->
<div class="container">
  <form action="#ready" method="post">
   <div class="row">
      <div class="col-25">
        <label for="lname">Фамилия</label>
      </div>
      <div class="col-75">
        <input type="text" id="lname" name="lastname" placeholder="Ваша фамилия">
      </div>
    </div>
   <div class="row">
      <div class="col-25">
        <label for="fname">Имя</label>
      </div>
      <div class="col-75">
        <input type="text" id="fname" name="firstname" placeholder="Ваше имя">
      </div>
    </div>
    
	<div class="row">
      <div class="col-25">
        <label for="fname">Отчество</label>
      </div>
      <div class="col-75">
        <input type="text" id="fname" name="midname" placeholder="Ваше отчество">
      </div>
    </div>
    <!-- <div class="row">
      <div class="col-25">
        <label for="subject">Комментарий</label>
      </div>
      <div class="col-75">
        <textarea id="subject" name="subject" placeholder="" style="height:200px"></textarea>
      </div>
    </div> -->
    <div class="row">
      <input type="submit" value="Отправить">
    </div>
  </form>
</div>   

 </body>

<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

error_reporting(0);
$firstname = $_POST['firstname'];
$lastname = $_POST['lastname'];
$midname = $_POST['midname'];

if ($firstname != ""){

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("list.xlsx");
$sheet = $spreadsheet->getActiveSheet();

$r = $lastname . " " . $firstname . " " . $midname;
$counterFirst = $counterFirst+1;
$r = iconv( "UTF-8", "cp1251",  $r );

$highestRow = $sheet->getHighestRow();
$highestRow1 = $highestRow + 1;
$sheet->setCellValue('A'.$highestRow1, $lastname);
$sheet->setCellValue('B'.$highestRow1, $firstname);
$sheet->setCellValue('C'.$highestRow1, $midname);

$writer = new Xlsx($spreadsheet);
$writer->save('list.xlsx');

$e = fopen('list.html', 'a+'); // Открываем файл
fwrite($e, "<table><tr><td>".$r."</td></tr></table>"); // Записываем данные
fclose($e);
} 
else {} // Закрываем файл
?>