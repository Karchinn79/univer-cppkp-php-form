
 <head>
 <br>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>Очистить список</title>
  <link rel="stylesheet" type="text/css" href="./css/styleclear.css">
  <center><h3>Вы уверены, что хотите очистить список?<br>Перед этим рекомендуется скачать существующий excel файл</h3><br>
  <a href="./list.xlsx" target="_blank">Скачать excel файл</a>
  <br>
  <a href="./list.html" target="_blank">Посмотреть список</a></center>
 </head>
 <body><br><br>
<center>
  <form action="#ready" method="post">
   <div class="row">
   
        <center>Для очистки списка введите в это поле "Сбросить"</center>

      <div class="col-25">
        <input type="text" id="lname" name="sure" placeholder="Сбросить">
      </div>
    </div><br>
	<div class="row">
      <input type="submit" value="Отправить">
    </div>
  </form>

</center>
 </body>

<?php
error_reporting(0);
$sr = $_POST['sure'];

$r="<?php session_cache_limiter('nocache'); ?> <head>
  <a href=\"./clear.php\" target=\"_blank\">Очистить список</a><br><br>
  <a href=\"./unik.php\" target=\"_blank\">Добавить имя</a><br>
  <a href=\"./list.xlsx\" target=\"_blank\">Скачать excel файл</a><br><br>
	</head>";

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
if ($sr == "сбросить"){goto das;}
if ($sr == "Сбросить"){
das:
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("list.xlsx");
$sheet = $spreadsheet->getActiveSheet();
$sheet->removeColumn('A');
$sheet->removeColumn('B');
$sheet->removeColumn('C');
gc_collect_cycles();

$filename = "list.html";
$handle = fopen($filename, 'r+');
ftruncate($handle, 0);
$e = fopen('list.html', 'a+'); // Открываем файл
$r = iconv( "UTF-8", "cp1251",  $r );
fwrite($e, $r);
fclose($e);
 // Закрываем файл

$sheet->removeColumn('A');
$writer = new Xlsx($spreadsheet);
$writer->save('list.xlsx');
echo "Список успешно сброшен";
}
else {
	echo "Введите подтверждение";
}
?>