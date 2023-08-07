<?php
require 'C:\php-8.2.8-Win32-vs16-x64\vendor\autoload.php'; // Include the PhpSpreadsheet autoloader

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

// Function to sanitize user input to prevent XSS and other attacks
function sanitizeInput($data)
{
    return htmlspecialchars(trim($data), ENT_QUOTES, 'UTF-8');
}

if ($_SERVER["REQUEST_METHOD"] === "POST") {
    // Sanitize the form data
    $paper_title = sanitizeInput($_POST["paper_title"]);
    $authors = sanitizeInput($_POST["authors"]);
    $year = intval($_POST["year"]);
    $award_category = sanitizeInput($_POST["award_category"]);
    $doi_or_pdf = sanitizeInput($_POST["doi_or_pdf"]);
    $update_authors = sanitizeInput($_POST["update_authors"]);
    $impact_significance = sanitizeInput($_POST["impact_significance"]);
    $proposed_citation = sanitizeInput($_POST["proposed_citation"]);
    $nominator_name = sanitizeInput($_POST["nominator_name"]);
    $relationship_to_authors = sanitizeInput($_POST["relationship_to_authors"]);
    $nominator_address = sanitizeInput($_POST["nominator_address"]);
    $nominator_country = sanitizeInput($_POST["nominator_country"]);


    // Output the file to the user for download
    //header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //header('Content-Disposition: attachment;filename="' . $filename . '"');
    //header('Cache-Control: max-age=0');
    //readfile($filename);
    //unlink($filename); // Delete the temporary file after download

    exit;
}
