<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

$host = 'localhost';
$db   = 'technical_assessment';
$user = 'root';
$pass = '';

$conn = new mysqli($host, $user, $pass, $db);
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

$name = trim($_POST['name']);
$email = filter_var(trim($_POST['email']), FILTER_VALIDATE_EMAIL);
$age = filter_var($_POST['age'], FILTER_VALIDATE_INT);

if (!$name || !$email || !$age) {
    die("All fields are required and must be valid.");
}

if (!isset($_FILES['excel_file']) || $_FILES['excel_file']['error'] != UPLOAD_ERR_OK) {
    die("File upload error.");
}

$fileInfo = pathinfo($_FILES['excel_file']['name']);
$allowedExtensions = ['xls', 'xlsx'];
if (!in_array(strtolower($fileInfo['extension']), $allowedExtensions)) {
    echo '<div style="color: red; font-size: 18px; text-align: center; padding: 20px;">Invalid file format. Only Excel files are allowed.</div>';
    exit();
}

$stmt = $conn->prepare("INSERT INTO users (name, email, age) VALUES (?, ?, ?)");
$stmt->bind_param('ssi', $name, $email, $age);
$stmt->execute();
$user_id = $stmt->insert_id;
$stmt->close();

$targetDir = 'uploads/';
$targetPath = $targetDir . basename($_FILES['excel_file']['name']);

if (!is_dir($targetDir)) {
    mkdir($targetDir, 0777, true);
}

if (!move_uploaded_file($_FILES['excel_file']['tmp_name'], $targetPath)) {
    die("Failed to move uploaded file.");
}

try {
    $spreadsheet = IOFactory::load($targetPath);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
    die('Error loading file: ' . $e->getMessage());
}

$worksheet = $spreadsheet->getActiveSheet();
$rows = [];
$errorMessages = [];
$successCount = 0;

foreach ($worksheet->getRowIterator() as $rowIndex => $row) {
    if ($rowIndex == 1) continue;

    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(false);
    $rowData = [];

    foreach ($cellIterator as $cell) {
        $rowData[] = $cell->getValue();
    }

    list($productName, $quantity, $price) = $rowData;

    if (empty($productName)) {
        $errorMessages[] = "Row $rowIndex: Product name is missing.";
        continue;
    }
    if (!is_numeric($quantity) || $quantity <= 0) {
        $errorMessages[] = "Row $rowIndex: Quantity must be a valid number greater than 0.";
        continue;
    }
    if (!is_numeric($price) || $price <= 0) {
        $errorMessages[] = "Row $rowIndex: Price must be a valid number greater than 0.";
        continue;
    }

    $stmt = $conn->prepare("INSERT INTO products (product_name, quantity, price) VALUES (?, ?, ?)");
    $stmt->bind_param('sid', $productName, $quantity, $price);
    $stmt->execute();
    $successCount++;
}
$stmt->close();
$conn->close();

unlink($targetPath);

if ($successCount > 0) {
    echo '<div style="color: green; font-size: 18px; text-align: center; padding: 20px;">Form and Excel data submitted successfully! Processed ' . $successCount . ' product(s).</div>';
}

if (!empty($errorMessages)) {
    echo '<div style="color: red; font-size: 16px; text-align: center; padding: 20px;"><strong>Errors:</strong><br>' . implode('<br>', $errorMessages) . '</div>';
}
?>
