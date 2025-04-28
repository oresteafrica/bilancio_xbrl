<?php
//------------------------------------------------------------------------------
// importazione librerie
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Font; 
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Settings;
use PhpOffice\PhpSpreadsheet\Collection\Memory;
use PhpOffice\PhpSpreadsheet\Cache\MemoryCache;
//------------------------------------------------------------------------------
// identificazione temporale
$current_year = date('Y');
$currentMonth = date("m");
$previous_year = $current_year - 1;
$pre_previous_year = $current_year - 2;
if ($currentMonth == 12) { $anno_ora = $current_year; } else { $anno_ora = $current_year - 1; }
$anno_prima = $anno_ora - 1;
$dataBilancio = date_create($anno_ora.'-12-31');
$esRifIni = date_create($anno_ora.'-01-01');
$esRifEnd = date_create($anno_ora.'-12-31');
$esPreIni = date_create($anno_prima.'-01-01');
$esPreEnd = date_create($anno_prima.'-12-31');
$now = new DateTime();
$date = $now->format('dmy');
$time = $now->format('His');
$nowString = $date.'_'.$time;
//------------------------------------------------------------------------------
// file Excel di riferimento
//$xlsxBilancio = 'xlsx/bilancio_ordinario.xlsx';
$xlsxBilancio = 'xlsx/bilancio_abbreviato.xlsx';
$workbook = IOFactory::load($xlsxBilancio);
// inserimento dati da programma nella tab "indice"
$sheet = $workbook->getSheetByName('indice');
//$sheet->setCellValue('C3', $dataBilancio->format('d/m/Y'));
$sheet->setCellValue('G3', $esRifIni->format('d/m/Y'));
$sheet->setCellValue('I3', $esRifEnd->format('d/m/Y'));
$sheet->setCellValue('G4', $esPreIni->format('d/m/Y'));
$sheet->setCellValue('I4', $esPreEnd->format('d/m/Y'));
//------------------------------------------------------------------------------
// database e tabelle di riferimento
$host = 'localhost';
$dbname = 'wp-playground';
$username = 'oreste';
$password = 'vaffax';

try {
    $pdo = new PDO("mysql:host=$host;dbname=$dbname;charset=utf8mb4", $username, $password);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
} catch (PDOException $e) {
	file_put_contents($debugFile, "Errore di collegamento con la base dati: " . $e->getMessage() . PHP_EOL, FILE_APPEND);
    die("Non riesco ad aprire il database. " . $e->getMessage());
}

// tabelle
// $tBilancio = 'contab_personia_bilancio_ordinario_xlsx';
$tBilancio = 'contab_personia_bilancio_abbreviato_xlsx';

// Debug
$debugFile = 'debug/' . $nowString.'_debug.txt';
$debugHandle = fopen($debugFile, 'a');
$lineOnFile = '';
//------------------------------------------------------------------------------
// iterazione nella tabella
$sql = "SELECT tab, cell, val, txt FROM $tBilancio ORDER BY tab, cell";
$stmt = $pdo->prepare($sql);
$stmt->execute();
$rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

foreach ($rows as $row) {
	$tab = (string) $row['tab'];
	$cell = (string) $row['cell'];
	$val = (float) $row['val'];
	$txt = (string) $row['txt'];

	// Debug
	$lineOnFile .= str_repeat('-',60) . PHP_EOL;
	$lineOnFile .= "Tab: $tab, Cell: $cell, Value: $val, Text: $txt" . PHP_EOL;

	// inserimento dati nel file Excel
	$sheet = $workbook->getSheetByName($tab);
	if ($sheet === null) {
		throw new Exception($tab . ' non risulta in ' . $xlsxBilancio);
	}
	// criterio: se val = -1 vuol dire che il codice in txt si riferisce all'anno precedente

	// contiene un valore numerico e non contiene alcun testo
	if ($txt == '') {
		$sheet->setCellValueExplicit($cell, $val, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
		continue;
	}
	if ($tab == 'T0000') {
		if ( strlen($txt) > 1 ) {
			$sheet->setCellValueExplicit($cell, $txt, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
			continue;
		}
	} else {
		// contiene una formula
		if (substr($txt, 0, 1)==='=') {
			$sheet->setCellValueExplicit($cell,$txt,\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA);
			continue;
		}
		// le  seguenti condizioni si possono riassumere con $txt == '0', si mantengono separate al fine
		// di facilitare la programmazione futura
		// contiene la stringa 0 relativa all'anno contabile in corso
		if ($txt == '0' AND $val != -1) {
			$sheet->setCellValueExplicit($cell, 0, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
			continue;
		}
		// contiene la stringa 0 relativa all'anno contabile precedente a quello in corso
		if ($txt == '0' AND $val == -1) {
			$sheet->setCellValueExplicit($cell, 0, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
			continue;
		}
		// le  seguenti condizioni si possono riassumere con preg_match('/[0-9-]+/', $txt), si mantengono separate al fine
		// di facilitare la programmazione futura
		// contiene un codice contabile riferito all'anno contabile in corso
		if (preg_match('/[0-9-]+/', $txt) && $val != -1) {
			$totCode = calcolaCodici($pdo,$txt,$anno_ora,$lineOnFile);

			// Debug
			$lineOnFile .= 'preg_match(\'/[0-9-]+/\', $txt) && $val != -1)' . PHP_EOL;
			$lineOnFile .= '$totCode = ' . $totCode . PHP_EOL;

			$sheet->setCellValueExplicit($cell, $totCode, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
			continue;
		}
		// contiene un codice contabile riferito all'anno contabile precedente a quello in corso
		if (preg_match('/[0-9-]+/', $txt) && $val == -1) {
			$totCode = calcolaCodici($pdo,$txt,$anno_prima);
			$sheet->setCellValueExplicit($cell, $totCode, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
			continue;
		}
		// contiene una stringa di testo utile
		if ( strlen($txt) > 1 AND ! preg_match('/[0-9-]+/', $txt) ) {
			$sheet->setCellValueExplicit($cell, $txt, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
			continue;
		}
	}
}
echo 'Elaborazione terminata, procedo con la scrittura dei file' . PHP_EOL;

// Debug
if ($debugHandle) {
	fwrite($debugHandle, $lineOnFile);
	fclose($debugHandle);
    chmod($debugFile, 0664);  // Makes file group-writable
    chown($debugFile, 'www-data');  // Requires sudo/root
	echo 'Debug file scritto' . PHP_EOL;
	echo $debugFile . PHP_EOL;
} else {
	echo 'Debug file NON scritto' . PHP_EOL;
}

// scrittura nuovo file Excel
try {
	$outputFileName = 'xlsx/' . $nowString.'.xlsx';
	$fileXlsx = $outputFileName;
	$writer = IOFactory::createWriter($workbook, 'Xlsx');
	$writer->save($fileXlsx);
	echo 'XLSX file scritto' . PHP_EOL;
	echo $fileXlsx . PHP_EOL;
} catch (\Exception $e) {
    echo "Errore scrittura nuovo file Excel: " . $e->getMessage();
}
//------------------------------------------------------------------------------
function calcolaCodici($pdo,$txt,$anno,&$lineOnFile = '') {
	$codes = explode(';', $txt);
	$codes = array_map('trim', $codes);

	// Debug
	$lineOnFile .= 'number of codes = ' . count($codes) . PHP_EOL;
	$lineOnFile .= '$codes = ' . print_r($codes, true) . PHP_EOL;

	$totCode = 0;
	foreach ($codes as $code) {
		$firstByte = substr($code, 0, 1);
		if ( $code === '0' || ( $firstByte !== 'n' && intval($code) == 0 ) ) { continue; }

		// occorre discriminare fra codici che dovranno restituire un valore negativo o positivo
		// si userà la enne (n) davanti un codice per indicarlo come negativo, a quel punto questa routine
		// sottrae invece di sommare

		$meno = false;
		if ($firstByte === 'n') {
			$meno = true;
			$code = substr($code, 1);
		}		

		try {
		    $sqlCodes = "SELECT SUM(m.Montante) AS total_sum
		    				FROM contab_personia_movimenti m 
		    				JOIN contab_personia_piano p ON m.Codice = p.Id 
		    				WHERE YEAR(m.Data) = :anno AND p.Codice LIKE :code";
			$stmt = $pdo->prepare($sqlCodes);
			$stmt->bindParam(':anno', $anno, PDO::PARAM_INT);
			$stmt->bindValue(':code', $code . '%', PDO::PARAM_STR);
			$stmt->execute();
			$result = $stmt->fetch(PDO::FETCH_ASSOC);
			$totalSum = $result ? ($result['total_sum'] !== null ? $result['total_sum'] : 0) : 0;
			if ($meno) { $totCode -= $totalSum; } else { $totCode += $totalSum; }
		} catch (PDOException $e) {
		    echo "Errore calcolo contabile: " . $e->getMessage();
		}

		// Debug
        $lineOnFile .= '$meno = ' . ($meno ? 'true' : 'false') . ' | ';
        $lineOnFile .= '$firstByte = ' . $firstByte . ' | ';
        $lineOnFile .= '$code = ' . $code . ' | ';
        $lineOnFile .= '$totalSum = ' . $totalSum . PHP_EOL;

	}
	// return $totCode;
	return (int) round($totCode);
}
//------------------------------------------------------------------------------
?>
