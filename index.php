 <?php

 require_once 'vendor/autoload.php'; 

 use PhpOffice\PhpWord\PhpWord; 
 use PhpOffice\PhpWord\IOFactory; 

// New instance
 $phpWord = new PhpWord(); 

// New section
 $section = $phpWord->addSection();

// Functions to generate filename correctly
function removeSpaces($palavra){
 	$palavra = trim($palavra);
 	$palavra = str_replace('  ', '', $palavra);
 	return $palavra;
}

function treatLetters($info){
 	$vogais = array('À', 'à', 'Á', 'á', 'Â', 'â', 'Ã', 'ã','Ç', 'ç','È', 'è', 'É', 'é', 'Ê', 'ê','Ì', 'ì', 'Í', 'í', 'Î', 'î', 'Ñ', 'ñ','Ò', 'ò', 'Ó', 'ó', 'Ô', 'ô', 'Õ', 'õ', 'Ù', 'ù', 'Ú', 'ú', 'Û', 'û','Ý', 'ý');
 	$ent = array('a', 'a', 'a', 'a', 'a', 'a', 'a', 'a','c', 'c','e', 'e', 'e', 'e', 'e', 'e','i', 'i', 'i', 'i', 'i', 'i','n', 'n','o', 'o', 'o', 'o', 'o', 'o', 'o', 'o','u', 'u', 'u', 'u', 'u', 'u','y', 'y');
 	return str_replace( $vogais, $ent, $info);
}

function generateFilename($palavra){
	//Retira preposições curtas
 	$palavra = trim($palavra);
 	$palavra = strtolower($palavra);
 	$retira = array('  ',' a ',' e ',' o ',' na ',' no ',' em ',' com ',' que ',' de ',' da ',' do ',' ',' - ');
 	$substitui = array('','-','-','-','-','-','-','-','-','-','-','-','-','-');
 	$tempPalavra = str_replace( $retira, $substitui, $palavra);
 	$tempPalavra = treatLetters($tempPalavra);
 	return $tempPalavra;
}

function treatText($desc){
 	$desc = str_replace('&aacute;','á',$desc);
 	$desc = str_replace('&eacute;','é',$desc);
 	$desc = str_replace('&iacute;','í',$desc);
 	$desc = str_replace('&oacute;','ó',$desc);
 	$desc = str_replace('&uacute;','ú',$desc);
 	$desc = str_replace('&atilde;','ã',$desc);
 	$desc = str_replace('&otilde;','õ',$desc);
 	$desc = str_replace('&acirc;','â',$desc);
 	$desc = str_replace('&ecirc;','ê',$desc);
 	$desc = str_replace('&icirc;','î',$desc);
 	$desc = str_replace('&ocirc;','ô',$desc);
 	$desc = str_replace('&ucirc;','û',$desc);
 	$desc = str_replace('&ccedil;','ç',$desc);
 	$desc = str_replace('&Aacute;','Á',$desc);
 	$desc = str_replace('&Eacute;','É',$desc);
 	$desc = str_replace('&Iacute;','Í',$desc);
 	$desc = str_replace('&Oacute;','Ó',$desc);
 	$desc = str_replace('&Uacute;','Ú',$desc);
 	$desc = str_replace('&Atilde;','Ã',$desc);
 	$desc = str_replace('&Otilde;','Õ',$desc);
 	$desc = str_replace('&Acirc;','Â',$desc);
 	$desc = str_replace('&Ecirc;','Ê',$desc);
 	$desc = str_replace('&Icirc;','Î',$desc);
 	$desc = str_replace('&Ocirc;','Ô',$desc);
 	$desc = str_replace('&Ucirc;','Û',$desc);
 	$desc = str_replace('&Ccedil;','Ç',$desc);
 	$desc = str_replace('&agrave;','à',$desc);
 	$desc = str_replace('&Agrave;','À',$desc);
 	$desc = str_replace('&sup2;','²',$desc);
 	$desc = str_replace('&ldquo;','“',$desc);
 	$desc = str_replace('&rdquo;','”',$desc);
 	$desc = str_replace('&quot;','',$desc);
 	$desc = str_replace('&nbsp;',' ',$desc);
 	$desc = str_replace('\"','"', $desc);
 	$desc = str_replace('<a href=','<a title="" href=', $desc);
 	$desc = preg_replace("/style=\"(.+?)\"/", "", $desc);
 	$desc = str_replace('<ul>','<ul class="topicos">', $desc);
 	$desc = str_replace('<p> </p>','', $desc);

 	return($desc);
}



// Get all doc files
 $docs = 'docs/*.docx';
 // Counter
 $cont = 0;

 foreach(glob($docs) as $doc):


 	$filename = basename($doc);

 	// Remove Dots and numbers
 	$nomepagina = str_replace( $cont. '. ', '', $filename);
 	$nomepagina = str_replace( $cont. ' ', '', $filename);

 	$nomepagina = removeSpaces($nomepagina);
 	$nomepagina = treatLetters($nomepagina);
 	$nomepagina = generateFilename($nomepagina);
 	$nomepagina = treatText($nomepagina);

 	// Generate filename before conversion
 	$filename = str_replace('.docx', '.php', $filename);

 	$phpWord = IOFactory::load($doc);

 	$cont++;

 	// Do the conversion
 	$htmlWriter = new \PhpOffice\PhpWord\Writer\HTML($phpWord);

	 // Directory to save those converted files
 	$htmlWriter->save('files-converted/'.$filename);

 endforeach;
 echo $cont . ' file(s) to convert';


 ?>