<?php
error_reporting(E_ALL ^ E_NOTICE);
include 'excel_reader.php';

$orden = new excel_reader('listado_sincedula.xls');
$orden->setColumnName(1);
$orden->setOutputEncoding('CP1251');

#ejecutamos la lectura del archivo excel:
$orden->read();

foreach($orden as $sheets) {
	echo "Hoja &lt <br />";
	$columnas = $orden->columns();
	$filas = $orden->numRows();
	for ($i = 2; $i <= $filas; $i++) {
		echo "Fila &lt <br />";
		$fila = $orden->rows($i);
		foreach($fila as $k=>$v) {
			print_r(nl2br("celda: {$k}= {$v}\n"));
		}
	}
	/* #forma alternativa
	 foreach($orden->rows() as $fila) {
		foreach($fila as $k=>$v) {
		print_r(nl2br("celda: {$k}= {$v}\n"));
		}
		}
		*/
}
?>