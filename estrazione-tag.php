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
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
//------------------------------------------------------------------------------
// identificazione temporale
$now = new DateTime();
//------------------------------------------------------------------------------
// costanti
$shortLine = str_repeat('-', 30) . PHP_EOL;
$longLine = str_repeat('-', 80) . PHP_EOL;
//------------------------------------------------------------------------------
// variabili
//------------------------------------------------------------------------------
// file di riferimento
// File Excel con codici e dati
$xlsxDataName = 'xlsx/200425_161333.xlsx';			// <-- sostituire con il nome del file generato da bilancio.php
//------------------------------------------------------------------------------
// Iterazione dei fogli nel file Excel

$spreadsheet = IOFactory::load($xlsxDataName);
$codeData = [];
foreach ($spreadsheet->getSheetNames() as $sheetName) {
    $sheet = $spreadsheet->getSheetByName($sheetName);
    // Determine the type of sheet
    $type = identifySheetType($sheet);
    if ($type === 'Monovalue') {
        // Monovalue sheet processing
        $code = $sheet->getCell('D3')->getCalculatedValue();
        $title = trim($sheet->getCell('B7')->getCalculatedValue());
        $value = $sheet->getCell('D11')->getCalculatedValue();
        if ($value !== null && $value !== '') { $codeData[] = [$code, $title, $value]; }
    } elseif ($type === 'Monodimensional') {
        // Monodimensional sheet processing
        $row = 10;
        while ($sheet->getCell("A{$row}")->getValue() !== null) {
            $code = $sheet->getCell("A{$row}")->getCalculatedValue();
            $title = trim($sheet->getCell("C{$row}")->getCalculatedValue());
            $value = $sheet->getCell("D{$row}")->getCalculatedValue() ?? ''; // Use empty string if value is missing
            if ($value !== null && $value !== '') { $codeData[] = [$code, $title, $value]; }
            $row++;
        }
    } elseif ($type === 'Multidimensional') {
        // Multidimensional sheet processing
        $colCodes = [];
        $rowCodes = [];
        $colTitles = [];
        $rowTitles = [];
        // Collect column codes and titles (starting at A11)
        $row = 11;
        while ($sheet->getCell("A{$row}")->getValue() !== null) {
            $colCodes[] = $sheet->getCell("A{$row}")->getCalculatedValue();
            $colTitles[] = trim($sheet->getCell("C{$row}")->getCalculatedValue());
            $row++;
        }
        // Collect row codes and titles (starting at D3)
        $colIndex = Coordinate::columnIndexFromString('D'); // Start from column D
        while ($sheet->getCell(Coordinate::stringFromColumnIndex($colIndex) . '3')->getValue() !== null) {
            $rowCodes[] = $sheet->getCell(Coordinate::stringFromColumnIndex($colIndex) . '3')->getCalculatedValue();
            $rowTitles[] = trim($sheet->getCell(Coordinate::stringFromColumnIndex($colIndex) . '10')->getCalculatedValue());
            $colIndex++;
        }
        // Iterate over column codes and row codes
        foreach ($colCodes as $i => $colCode) {
            foreach ($rowCodes as $j => $rowCode) {
                $combinedCode = $colCode . '_' . $rowCode;
                $combinedTitle = $colTitles[$i] . ' ' . $rowTitles[$j];

                // Calculate the intersection cell
                $columnLetter = Coordinate::stringFromColumnIndex(4 + $j); // Start from column D (index 4)
                $rowNumber = 11 + $i; // Start from row 11

                // Get the value at the intersection
                $cellCoordinate = $columnLetter . $rowNumber;
                $value = $sheet->getCell($cellCoordinate)->getCalculatedValue() ?? '';
                
                if ($value !== null && $value !== '') { $codeData[] = [$combinedCode, $combinedTitle, $value]; }
           }
        }
    }
}

writeArrayToCsv($codeData, 'xlsx');

// Qui $codeData è completo ora si procede nella ricerca del codice in mapping.xml

$xbrlCode = makeCSVfromXML($codeData);

// Qui i dati sono sufficienti per costruire il file XBRL

$xbrlCodeNoDupl = filterUniqueBySecondItem($xbrlCode);
$xbrlBody = makeXbrlArray($xbrlCodeNoDupl);
arrayToTextFile($xbrlBody, 'bilancio', 'xbrl');
echo $longLine . 'Programma concluso' . PHP_EOL . $longLine;

//----------------------------------------------------------------------------------------------------------------------------------
/*
Restituisce le linee con valori unici del secondo componente della linea
*/
function filterUniqueBySecondItem(array $lines): array {
    $uniqueLines = [];
    $seenItems = [];
    foreach ($lines as $line) {
        $secondItem = $line[1];
        // If we haven't seen this second item before, add it to results
		if ( in_array($secondItem,$seenItems) ) {
			continue;
		} else {
		    $uniqueLines[] = $line;
		    $seenItems[] = $secondItem;
		}
    }
    return $uniqueLines;
}
//----------------------------------------------------------------------------------------------------------------------------------
/*
Costruisci una matrice che contiene tutte le informazioni necessarie per generare il bilancio su file xbrl
Unique values:
$xbrlCode[0][2]: itcc-ci
$xbrlCode[0][3]: string, monetary, boolean, decimal, shares, date
$xbrlCode[0][4]: instant, duration
$xbrlCode[0][5]: valore stringa o numero
esempio:
0	T0002.D01.1.001.002.004.003.000,
1	ImmobilizzazioniMaterialiDestinateAllaVendita,
2	itcc-ci,
3	monetary,
4	instant,
5	0

la chiusura è statica
</xbrl>
*/

function makeXbrlArray($csvArray) {
	$xbrlBody = [];
	$now = new DateTime();
	$xbrlBody[] = file_get_contents('txt/xbrl_01.txt');
	$xbrlBody[] = '<!-- Version' . $now->format('Y-m-d H:i:s') . '-->';
	$xbrlBody[] = file_get_contents('txt/xbrl_03.txt');
	$xbrlBody[] = file_get_contents('txt/xbrl_04.txt');
	foreach ($csvArray as $xbrlLine) {

		// linee che non riesco a definire
		if ( $xbrlLine[3] == 'string' && $xbrlLine[4] == 'duration' ) { continue; }
		if ( $xbrlLine[3] == 'shares' && $xbrlLine[4] == 'duration' ) { continue; }
		if ( $xbrlLine[3] == 'shares' && $xbrlLine[4] == 'instant' ) { continue; }

		$xbrlCompleteLine = '';
		$tag = $xbrlLine[2].':'.$xbrlLine[1];

		// i = instant , d = duration
		if ( $xbrlLine[4] == 'duration' ) { $cref = 'contextRef="c2024_d"'; } else { $cref = 'contextRef="c2024_i"'; }

		switch ($xbrlLine[3]) {
		    case 'string':
					$xbrlCompleteLine = '<'.$tag.' ' . $cref . '>'.$xbrlLine[5].'</'.$tag.'>';
		        break;
		    case 'textBlock':
					$xbrlCompleteLine = '<'.$tag.' ' . $cref . '>'.$xbrlLine[5].'</'.$tag.'>';
		        break;
		    case 'nonnum:textBlock':
					$xbrlCompleteLine = '<'.$tag.' ' . $cref . '>'.$xbrlLine[5].'</'.$tag.'>';
		        break;
		    case 'monetary':
				$xbrlCompleteLine = '<'.$tag.' ' . $cref . ' unitRef="EUR" decimals="0">'.$xbrlLine[5].'</'.$tag.'>';
		        break;
		    case 'boolean':
		    	if ( strtolower($xbrlLine[5]) == 'si' || $xbrlLine[5] == 1 || $xbrlLine[5] == '1' ) {
					$xbrlCompleteLine = '<'.$tag.' ' . $cref . '>true</'.$tag.'>';
		    	} elseif ( strtolower($xbrlLine[5]) == 'no' || $xbrlLine[5] == 0 || $xbrlLine[5] == '0' ) {
					$xbrlCompleteLine = '<'.$tag.' ' . $cref . '>false</'.$tag.'>';
		    	} else {
					$xbrlCompleteLine = '<'.$tag.' ' . $cref . '>'.$xbrlLine[5].'</'.$tag.'>';
		    	}
		        break;
		    case 'decimal':
				$xbrlCompleteLine = '<'.$tag.' ' . $cref . ' unitRef="EUR" decimals="0">'.$xbrlLine[5].'</'.$tag.'>';
		        break;
		    case 'shares':
				$xbrlCompleteLine = '<'.$tag.' ' . $cref . ' unitRef="EUR" decimals="0">'.$xbrlLine[5].'</'.$tag.'>';
		        break;
		    case 'date':
				$dateISO = convertItalianDateToISO($xbrlLine[5]);
				$xbrlCompleteLine = '<'.$tag.' ' . $cref . '>'.$dateISO.'</'.$tag.'>';
		        break;
		    default:
		}
		$xbrlBody[] = $xbrlCompleteLine;
	}
	$xbrlBody[] = '</xbrl>';
	return $xbrlBody;
}
//----------------------------------------------------------------------------------------------------------------------------------
/*
Check if the input string matches the Italian date format (dd/mm/yyyy)
Validate the date
Create ISO format (YYYY-MM-DD)
Return last day of 2024 if the format is invalid
*/
function convertItalianDateToISO($dateString) {
    if (preg_match('/^(\d{2})\/(\d{2})\/(\d{4})$/', $dateString, $matches)) {
        $day = $matches[1];
        $month = $matches[2];
        $year = $matches[3];
        if (checkdate($month, $day, $year)) {
            return sprintf('%04d-%02d-%02d', $year, $month, $day);
        }
    }
    return '2024-12-31';
}
//----------------------------------------------------------------------------------------------------------------------------------
/*
Scrivi file di testo da una matrice
Array to text file
*/
function arrayToTextFile($array, $prefix, $extension) {
	$now = new DateTime();
	$nowString = $now->format('dmy-His');
	$filename = 'xbrl/' . $prefix . '_' . $nowString . '.' . $extension;

	$txtname = 'xbrl/' . $prefix . '_errori_' . $nowString . '.txt';

	$result = file_put_contents($filename, implode(PHP_EOL, $array));

	if ($result) {

		$ftxt = fopen($txtname, 'w');
		fclose($ftxt);

	    chmod($filename, 0664);  // Makes file group-writable
	    chown($filename, 'www-data');  // Requires sudo/root (not recommended)

	    chmod($txtname, 0664);  // Makes file group-writable
	    chown($txtname, 'www-data');  // Requires sudo/root (not recommended)

	    echo $filename . ' scritto su disco' . PHP_EOL;

	} else {
		echo 'Non sono riuscito a scrivere ' . $filename . PHP_EOL;
	}
}
//----------------------------------------------------------------------------------------------------------------------------------
/*
Costruzione del CSV a partire dal file mapping.xml
*/
function makeCSVfromXML($codeData) {
	static $xmlMapName = 'xml/mapping.xml';
	$xbrlCode = [];
	foreach ($codeData as $codeLine) {
		$xmlCellAttr =  getXbrlAttributes($codeLine[0], $xmlMapName);
		if ($xmlCellAttr === null) { 
			// echo $codeLine[0] . ' non trovato' . PHP_EOL;
			continue;
		}
		array_unshift($xmlCellAttr, $codeLine[0]);
		array_push($xmlCellAttr, $codeLine[2]);
		$xbrlCode[] = $xmlCellAttr;
	}
	writeArrayToCsv($xbrlCode, 'xbrl');
	return $xbrlCode;
}
//----------------------------------------------------------------------------------------------------------------------------------
/*
Loads the XML file using SimpleXML
Registers the XBRL namespace to properly handle xbrl-prefixed attributes
Searches for a cell with the given code using XPath
If found, extracts the xbrl attributes (name, prefix, type, periodType)
Returns them as an associative array
*/
function getXbrlAttributes($cellCode, $xmlFile) {
    // Load the XML file
    $xml = simplexml_load_file($xmlFile);
    if ($xml === false) {
        throw new Exception("Failed to load XML file: $xmlFile");
    }

    // Register the xbrl namespace to access attributes with xbrl prefix
    $xml->registerXPathNamespace('xbrl', 'http://www.xbrl.org');

    // Search for the cell with the given code
    $cells = $xml->xpath("//cell[@code='$cellCode']");
    
    if (empty($cells)) {
        return null; // Cell code not found
    }

    $cell = $cells[0];
    
    // Get the xbrl attributes
	/*
    $attributes = [
        'name' => (string)$cell->attributes('xbrl', true)->name,
        'prefix' => (string)$cell->attributes('xbrl', true)->prefix,
        'type' => (string)$cell->attributes('xbrl', true)->type,
        'periodType' => (string)$cell->attributes('xbrl', true)->periodType
    ];
    */

	$attr = [
				(string)$cell->attributes('xbrl', true)->name,
				(string)$cell->attributes('xbrl', true)->prefix,
				(string)$cell->attributes('xbrl', true)->type,
				(string)$cell->attributes('xbrl', true)->periodType
			];


    return $attr;
}
//----------------------------------------------------------------------------------------------------------------------------------
/*
Writes an array to a CSV file.
@param array $data The array to write to the CSV file.
@param string $filename The name of the output CSV file.
*/
function writeArrayToCsv($data, $prefix) {
	$now = new DateTime();
	$nowString = $now->format('dmy-His');
	$filename = 'csv/' . $prefix . '_' . $nowString . '.csv';
    $file = fopen($filename, 'w');
    foreach ($data as $row) {
        fputcsv($file, $row);
    }
    fclose($file);
    chmod($filename, 0664);  // Makes file group-writable
    chown($filename, 'www-data');  // Requires sudo/root (not recommended)
    echo $filename . ' scritto su disco' . PHP_EOL;
}
//----------------------------------------------------------------------------------------------------------------------------------
/*
Identifies the type of sheet based on its structure.
@param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet
@return string The type of sheet: 'Monovalue', 'Monodimensional', or 'Multidimensional'.
*/
function identifySheetType($sheet) {
    // Check for Monovalue sheet
    if (
        $sheet->getCell('D3')->getValue() !== null &&
        $sheet->getCell('B7')->getValue() !== null &&
        $sheet->getCell('D10')->getValue() !== null &&
        $sheet->getCell('A11')->getValue() === null &&
        $sheet->getCell('A12')->getValue() === null
    	) {
        return 'Monovalue';
    }

    // Check for Monodimensional sheet
    $row = 10;
    if ($sheet->getCell("A{$row}")->getValue() !== null) {
        // Check if there are codes in column A and corresponding titles in column C
        while ($sheet->getCell("A{$row}")->getValue() !== null) {
            if ($sheet->getCell("C{$row}")->getValue() === null) {
                break; // Titles are missing, so it's not Monodimensional
            }
            $row++;
        }
        return 'Monodimensional';
    }

    // Check for Multidimensional sheet
    if (
        $sheet->getCell('A11')->getValue() !== null &&
        $sheet->getCell('C11')->getValue() !== null &&
        $sheet->getCell('D3')->getValue() !== null &&
        $sheet->getCell('D10')->getValue() !== null
    ) {
        return 'Multidimensional';
    }

    // Default fallback (if no type matches)
    return 'Unknown';
}
//----------------------------------------------------------------------------------------------------------------------------------
/*
 * Reads a CSV file and stores its contents in a PHP array
 * 
 * @param string $filePath Path to the CSV file
 * @param bool $hasHeaders Whether the CSV has a header row (default: false)
 * @param string $delimiter Field delimiter (default: ',')
 * @param string $enclosure Field enclosure (default: '"')
 * @param string $escape Escape character (default: '\\')
 * 
 * @return array Associative array if $hasHeaders is true, numeric array otherwise
 * @throws Exception If file cannot be opened
*/
function csvToArray(
    string $filePath, 
    bool $hasHeaders = false, 
    string $delimiter = ',', 
    string $enclosure = '"', 
    string $escape = '\\'
): array {
    // Check if file exists and is readable
    if (!file_exists($filePath) || !is_readable($filePath)) {
        throw new Exception("File not found or not readable: " . $filePath);
    }

    // Open the file for reading
    $file = fopen($filePath, 'r');
    if ($file === false) {
        throw new Exception("Failed to open file: " . $filePath);
    }

    $data = [];
    $headers = [];

    // Read headers if they exist
    if ($hasHeaders) {
        $headers = fgetcsv($file, 0, $delimiter, $enclosure, $escape);
        if ($headers === false) {
            fclose($file);
            throw new Exception("Failed to read headers from CSV");
        }
    }

    // Read the remaining data
    while (($row = fgetcsv($file, 0, $delimiter, $enclosure, $escape)) !== false) {
        // Skip empty rows
        if ($row === [null]) {
            continue;
        }

        // Create associative array if headers exist
        if ($hasHeaders) {
            $assocRow = [];
            foreach ($headers as $index => $header) {
                $assocRow[$header] = $row[$index] ?? null;
            }
            $data[] = $assocRow;
        } else {
            $data[] = $row;
        }
    }

    fclose($file);
    return $data;
}
//----------------------------------------------------------------------------------------------------------------------------------
?>
