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

    // Create a new Spreadsheet object
    // $spreadsheet = getActiveSpreadsheet().getSheetByName('https://docs.google.com/spreadsheets/d/18zU3HYqH6pu16hC56w69Fm9fe4LvDSX5UvysRFRO5hk/edit?usp=sharing');
    // $spreadsheet = new Spreadsheet();
    $spreadsheet = IOFactory::load('https://docs.google.com/spreadsheets/d/18zU3HYqH6pu16hC56w69Fm9fe4LvDSX5UvysRFRO5hk/edit?usp=sharing');
    $sheet = $spreadsheet->getActiveSheet(); 

    // Set the headers for the spreadsheet
    //$sheet->setCellValue('A1', 'Paper Title');
    //$sheet->setCellValue('B1', 'Authors');
    //$sheet->setCellValue('C1', 'Year');
    //$sheet->setCellValue('D1', 'Award Category');
    //$sheet->setCellValue('E1', 'DOI or PDF link');
    //$sheet->setCellValue('F1', 'Update on author institutions');
    //$sheet->setCellValue('G1', 'Impact and Significance');
    //$sheet->setCellValue('H1', 'Proposed Citation');
    //$sheet->setCellValue('I1', 'Nominator Name');
    //$sheet->setCellValue('J1', 'Relationship to Authors');
    //$sheet->setCellValue('K1', 'Address');
    //$sheet->setCellValue('L1', 'Country');

// Get the last row number (assuming data starts from row 2)
    $lastRow = $sheet->getHighestRow() + 1;
    
    // Set the data from the form
    $sheet->setCellValue('A' . $lastRow, $paper_title);
    $sheet->setCellValue('B' . $lastRow, $authors);
    $sheet->setCellValue('C' . $lastRow, $year);
    $sheet->setCellValue('D' . $lastRow, $award_category);
    $sheet->setCellValue('E' . $lastRow, $doi_or_pdf);
    $sheet->setCellValue('F' . $lastRow, $update_authors);
    $sheet->setCellValue('G' . $lastRow, $impact_significance);
    $sheet->setCellValue('H' . $lastRow, $proposed_citation);
    $sheet->setCellValue('I' . $lastRow, $nominator_name);
    $sheet->setCellValue('J' . $lastRow, $relationship_to_authors);
    $sheet->setCellValue('K' . $lastRow, $nominator_address);
    $sheet->setCellValue('L' . $lastRow, $nominator_country);

    // Save the spreadsheet to a file
    // $writer = new Xlsx($spreadsheet);
    // $filename = 'award_nomination_' . date('Ymd_His') . '.xlsx';
    // $writer->save($filename);
    // Save the modified spreadsheet
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save('https://docs.google.com/spreadsheets/d/18zU3HYqH6pu16hC56w69Fm9fe4LvDSX5UvysRFRO5hk/edit?usp=sharing');

    // Output the file to the user for download
    //header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //header('Content-Disposition: attachment;filename="' . $filename . '"');
    //header('Cache-Control: max-age=0');
    //readfile($filename);
    //unlink($filename); // Delete the temporary file after download

    exit;
}
