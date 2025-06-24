<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

function cleanString($str) {
    return trim(preg_replace('/\s+/', ' ', $str));
}

function isValidPostcode($value) {
    return preg_match("/^[1-9][0-9]{3}\s?[A-Z]{2}\$/", $value);
}

function isValidTime($value) {
    return (bool)DateTime::createFromFormat('H:i', trim($value)) ||
           (bool)DateTime::createFromFormat('H:i:s', trim($value));
}

function isValidDate($value) {
    return (bool)DateTime::createFromFormat('d-m-Y', trim($value)) ||
           (bool)DateTime::createFromFormat('Y-m-d', trim($value)) ||
           (bool)DateTime::createFromFormat('d/m/Y', trim($value));
}

function isValidGPS($value) {
    $value = str_replace(',', '.', trim($value));
    return is_numeric($value) && $value >= -180 && $value <= 180;
}

$fouten = [];

if ($_SERVER["REQUEST_METHOD"] === "POST" && isset($_FILES["excel_file"])) {
    $file = $_FILES["excel_file"]["tmp_name"];
    $spreadsheet = IOFactory::load($file);

    $sheetnames = $spreadsheet->getSheetNames();
    $selectedSheet = $_POST['sheet_name'] ?? $sheetnames[0];
    $sheet = $spreadsheet->getSheetByName($selectedSheet);

    $rows = $sheet->toArray();
    $headers = $rows[0];

    echo "<h3>Foutenoverzicht</h3>";
    echo "<table border='1'><tr><th>Rij</th><th>Kolom</th><th>Waarde</th><th>Fouttype</th></tr>";

    foreach ($rows as $i => $row) {
        if ($i === 0) continue;
        foreach ($row as $j => $cell) {
            $colName = $headers[$j] ?? "Kol$j";
            $waarde = is_string($cell) ? cleanString($cell) : $cell;

            if (is_string($cell)) {
                if ($waarde !== ltrim($cell)) {
                    $fouten[] = ['rij' => $i+1, 'kolom' => $colName, 'waarde' => $cell, 'type' => 'Leading whitespace'];
                }
                if ($waarde !== rtrim($cell)) {
                    $fouten[] = ['rij' => $i+1, 'kolom' => $colName, 'waarde' => $cell, 'type' => 'Trailing whitespace'];
                }
            }

            if (stripos($colName, 'postcode') !== false) {
                $delen = preg_split("/[ ,;|]+/", $waarde);
                foreach ($delen as $deel) {
                    if ($deel && !isValidPostcode($deel)) {
                        $fouten[] = ['rij' => $i+1, 'kolom' => $colName, 'waarde' => $deel, 'type' => 'Ongeldige postcode'];
                    }
                }
            }

            if (stripos($colName, 'tijd') !== false || stripos($colName, 'time') !== false) {
                if ($waarde && !isValidTime($waarde)) {
                    $fouten[] = ['rij' => $i+1, 'kolom' => $colName, 'waarde' => $waarde, 'type' => 'Ongeldige tijd'];
                }
            }

            if (stripos($colName, 'datum') !== false || stripos($colName, 'date') !== false) {
                if ($waarde && !isValidDate($waarde)) {
                    $fouten[] = ['rij' => $i+1, 'kolom' => $colName, 'waarde' => $waarde, 'type' => 'Ongeldige datum'];
                }
            }

            if (preg_match("/(lat|long)/i", $colName)) {
                if ($waarde === null || $waarde === '') continue;
                if (!isValidGPS($waarde)) {
                    $fouten[] = ['rij' => $i+1, 'kolom' => $colName, 'waarde' => $waarde, 'type' => 'Ongeldige GPS'];
                }
            }
        }
    }

    foreach ($fouten as $f) {
        echo "<tr><td>{$f['rij']}</td><td>{$f['kolom']}</td><td>{$f['waarde']}</td><td>{$f['type']}</td></tr>";
    }

    echo "</table>";

    // Samenvatting
    $tellingen = [];
    foreach ($fouten as $f) {
        $tellingen[$f['type']] = ($tellingen[$f['type']] ?? 0) + 1;
    }

    echo "<h3>Fouttellingen</h3><table border='1'><tr><th>Fouttype</th><th>Aantal</th></tr>";
    foreach ($tellingen as $type => $aantal) {
        echo "<tr><td>$type</td><td>$aantal</td></tr>";
    }
    echo "</table>";
}
?>

<!DOCTYPE html>
<html>
<head><title>Excel Validatie</title></head>
<body>
<h2>Upload een Excel bestand (.xlsx)</h2>
<form method="post" enctype="multipart/form-data">
  <label>Excel bestand:</label><br>
  <input type="file" name="excel_file" accept=".xlsx" required><br><br>

  <?php
  if (isset($sheetnames)) {
      echo "<label>Tabblad:</label><br>";
      echo "<select name='sheet_name'>";
      foreach ($sheetnames as $sheet) {
          $sel = ($sheet === $selectedSheet) ? "selected" : "";
          echo "<option $sel>$sheet</option>";
      }
      echo "</select><br><br>";
  }
  ?>

  <button type="submit">Valideer</button>
</form>
</body>
</html>
