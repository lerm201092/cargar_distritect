<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>Leer Archivo Excel</h1>
<?php
require_once 'PHPExcel/Classes/PHPExcel.php';
$archivo = "libro1.xlsx";
$inputFileType = PHPExcel_IOFactory::identify($archivo);
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
$objPHPExcel = $objReader->load($archivo);
$sheet = $objPHPExcel->getSheet(0); 
$highestRow = $sheet->getHighestRow(); 
$highestColumn = $sheet->getHighestColumn();
$insert="Delete From Productos; <br><br>";
for ($row = 2; $row <= $highestRow; $row++){ 
		$fila_a = $sheet->getCell("A".$row)->getValue();
		$fila_b = $sheet->getCell("B".$row)->getValue();
		$fila_c = $sheet->getCell("C".$row)->getValue();
		$fila_d = $sheet->getCell("D".$row)->getValue();
		$fila_e = $sheet->getCell("E".$row)->getValue();
		$fila_f = $sheet->getCell("F".$row)->getValue();
		$fila_g = $sheet->getCell("G".$row)->getValue();
		$fila_h = $sheet->getCell("H".$row)->getValue();
		$fila_i = $sheet->getCell("I".$row)->getValue();
		$fila_j = $sheet->getCell("J".$row)->getValue();
		$fila_k = $sheet->getCell("K".$row)->getValue();
		$fila_l = $sheet->getCell("L".$row)->getValue();
		$fila_m = $sheet->getCell("M".$row)->getValue();
		$fila_n = $sheet->getCell("N".$row)->getValue();
		$fila_o = $sheet->getCell("O".$row)->getValue();

		
		$insert .= "INSERT INTO PRODUCTOS (PRO_REFERENCIA, 
									      PRO_DESCRIPCION, 
										  PRO_ACCESORIOS, 
										  PRO_UNIDAD, 
										  PRO_USO, 
										  PRO_CAPACIDAD, 
										  PRO_PESO, 
										  PRO_DIAMETRO_ROSCA, 
										  PRO_ALTURA_ROSCA, 
										  PRO_DIAMETRO, 
										  PRO_ALTURA, 
										  PRO_COLOR, 
										  PRO_MATERIAL, 
										  PRO_PRECIO_UNI, 
										  PRECIO_CONTENIDO)<br>";
		$insert .= "VALUES ('$fila_a', 
							'$fila_b', 
							'$fila_c', 
							'$fila_d', 
							'$fila_e', 
							'$fila_f', 
							'$fila_g', 
							'$fila_h', 
							'$fila_i', 
							'$fila_j', 
							'$fila_k', 
							'$fila_l', 
							'$fila_m', 
							'$fila_n', 
							'$fila_o');";
		$insert .="<br><br>";	
	}
echo $insert."<br><br>"."Select * from Productos;";
?>
</body>
</html>
