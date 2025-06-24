<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

function isValidPostcode($value) {
    return preg_match("/^[1-9][0-9]{3}\s?[A-Z]{2}\$/", trim($value));
}

if ($_SERVER["REQUEST_METHOD"] === "POST" && isset($_FILES["excel_file"])) {
    $file = $_FILES["excel_file"]["tmp_name"];
    $spreadsheet = IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();

    echo "<table border='1'><tr><th>Rij</th><th>Kolom</th><th>Waarde</th><th>Status</th></tr>";

    foreach ($sheet->getRowIterator() as $rowIndex => $row) {
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);
        foreach ($cellIterator as $cellIndex => $cell) {
            $value = $cell->getValue();
            if ($value && preg_match("/[0-9]{4}\s?[A-Z]{2}/", $value)) {
                $delen = preg_split("/[ ,;|]+/", trim($value));
                foreach ($delen as $deel) {
                    if (!isValidPostcode($deel)) {
                        echo "<tr><td>{$rowIndex}</td><td>{$cellIndex}</td><td>{$deel}</td><td>Ongeldig</td></tr>";
                    }
                }
            }
        }
    }

    echo "</table>";
}
?>

<!DOCTYPE html>
<html>
<head><title>Postcode Validator</title></head>
<body>
<h2>Upload een Excel bestand (.xlsx)</h2>
<form method="post" enctype="multipart/form-data">
  <input type="file" name="excel_file" accept=".xlsx" required>
  <button type="submit">Valideer</button>
</form>
</body>
</html>
