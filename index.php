<?php

require_once 'vendor/autoload.php'; 

use PhpOffice\PhpWord\PhpWord; 
use PhpOffice\PhpWord\IOFactory; 

// New instance
$phpWord = new PhpWord(); 

// New section
$section = $phpWord->addSection();


// Get the .docx file
$phpWord = \PhpOffice\PhpWord\IOFactory::load('documentwordname.docx');

// put all .docx file content to a variable
$htmlWriter = new \PhpOffice\PhpWord\Writer\HTML($phpWord);

// Save the content in a .php file, you can change the extension file too.
$htmlWriter->save('myfilename.php');

// if you want to show the content of .docx file, just uncomment below:
//echo $htmlWriter;


?>