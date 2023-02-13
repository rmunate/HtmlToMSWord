<?php

namespace Rmunate\HtmlToMsword;

/**
 * Clase para crear Documento MS Word a partir de HTML
 * --------------------------------------------
 * Desarrollado por: Raul Mauricio Uñate Castro
 * raulmauriciounate@gmail.com
 */

class HtmlToWord 
{

	/* Porpiedades del Objeto */
	private $doc_file_separator;
	private $directoryImages;
	private $deleteOldFiles;
	private $replacements;
	private $expiredFiles;
	private $fullPathFile;
	private $originImages;
	private $stylesCSS;
	private $fileRoute;
	private $nameFile;
	private $encoding;
	private $i_img;
	private $html;

	/* Constructor de Clase */
	public function __construct() {
		$this->doc_file_separator = 'doc_file_separator'; //Valor por defecto
		$this->deleteOldFiles = true; //Eliminacion Activada por defecto.
        $this->expiredFiles = 1; //En Horas Horas (Numeros Enteros)
		$this->replacements = []; //Vacio Obligatorio
        $this->stylesCSS = ''; //Vacio Obligatorio
		$this->encoding = 'UTF-8'; //Valor por defecto
		$this->i_img = 0; //Valor por defecto
    }

	/* Tipo de Codificacion */
	public function encoding(string $enc = 'UTF-8'){
		$this->encoding = $enc;
	}

	/* Ruta del Archivo */
	public function fileRoute(string $route){

		/* Limpieza de Ruta */
		$route = $route .'/';
		$route = str_replace(array('//','\\'), '/', $route);

		/* Creacion de Directorio */
		if(!is_dir($route)){ 

			/* De no Existir crear el Directrorio */
			@mkdir($route, 0777); 
		}

		/* Registro de la Ruta De Salida */
		$this->fileRoute = $route;

	}

	/* Nombre del Documento */
	public function nameFile(string $name){

		/* Limpiar Nombre */
		$name = trim($name);
		
		//Reemplazamos la A y a
        $name = str_replace(
			array('Á', 'À', 'Â', 'Ä', 'á', 'à', 'ä', 'â', 'ª'),
			array('A', 'A', 'A', 'A', 'a', 'a', 'a', 'a', 'a'),
			$name
		);

		//Reemplazamos la E y e
		$name = str_replace(
			array('É', 'È', 'Ê', 'Ë', 'é', 'è', 'ë', 'ê'),
			array('E', 'E', 'E', 'E', 'e', 'e', 'e', 'e'),
			$name 
		);

		//Reemplazamos la I y i
		$name = str_replace(
			array('Í', 'Ì', 'Ï', 'Î', 'í', 'ì', 'ï', 'î'),
			array('I', 'I', 'I', 'I', 'i', 'i', 'i', 'i'),
			$name 
		);

		//Reemplazamos la O y o
		$name = str_replace(
			array('Ó', 'Ò', 'Ö', 'Ô', 'ó', 'ò', 'ö', 'ô'),
			array('O', 'O', 'O', 'O', 'o', 'o', 'o', 'o'),
			$name 
		);

		//Reemplazamos la U y u
		$name = str_replace(
			array('Ú', 'Ù', 'Û', 'Ü', 'ú', 'ù', 'ü', 'û'),
			array('U', 'U', 'U', 'U', 'u', 'u', 'u', 'u'),
			$name 
		);

		//Reemplazamos la N, n, C y c
		$name = str_replace(
			array('Ñ', 'ñ', 'Ç', 'ç'),
			array('N', 'n', 'C', 'c'),
			$name
		);

		//Reemplazamos Espacios
		$name = str_replace(' ', '_', $name);
		
		//Reemplazamos Caracteres Especiales
    	$name = preg_replace('([^A-Za-z0-9])', '', $name);

		/* Nombre Definitivo */
		// $name = strtoupper($name);

		/* Definir Ruta Completa Archivo */
		$route = $this->fileRoute . '/' . $name . '.doc';
		$fullPath = str_replace(array('//', '\\'), '/', $route);

		/* Guardar Nombre */
		$this->nameFile = $name;
		$this->fullPathFile = $fullPath;
	}

	/* Eliminar Archivos Viejos */
	public function deleteOldFiles(bool $swith = false, int $hours = 1){

		if ($swith) {

			$this->deleteOldFiles = true;
			$this->expiredFiles = $hours;

			/* Eliminar Imagenes */
			$files = glob($this->directoryImages. '*');
			foreach($files as $file){
				$lastModifiedTime = filemtime($file);
				$currentTime = time();
				$timeDiff = abs($currentTime - $lastModifiedTime)/(60*60);
				if(is_file($file) && $timeDiff > $this->expiredFiles){ 
					@unlink($file); 
				} 
			}

			/* Eliminar Documentos */
			$files = glob($this->fileRoute. '*');
			foreach($files as $file){
				$lastModifiedTime = filemtime($file);
				$currentTime = time();
				$timeDiff = abs($currentTime - $lastModifiedTime)/(60*60);
				if(is_file($file) && $timeDiff > $this->expiredFiles){ 
					@unlink($file); 
				} 
			}
			
			return true;

		} else {

			$this->deleteOldFiles = false;
			$this->expiredFiles = null;
			
			return false;
		}

	}

	/* Ruta donde se alojarán los SRC que se descarguen de internet (Solo Imagenes) */
	public function routeImages(string $routeImages){

		$route = $routeImages .'/';
		$route = str_replace(array('//','\\'), '/', $route);

		if(!is_dir($route)){ 
			@mkdir($route, 0777); 
		}

		$this->directoryImages = $route;
	}

	/* Origen de las Imagenes a Usar */
	public function originImages(string $origin = 'url'){
		$this->originImages = $origin;
	}

	/* Remplazos en el HTML */
	public function replacements(array $replacementsHTML){
		if (count($replacementsHTML) > 0) {
			$this->replacements = $replacementsHTML;
		}
	}

	/* Estilos para el documento */
	public function styles(array $arrayStyles){
		if (count($arrayStyles) > 0) {
			foreach ($arrayStyles as $key => $value) {
				$this->stylesCSS .= $key . '{' . $value . '} ';
			}
		}
	}

	/* Procesamiento o Descarga de Imagenes | Descarga de imagenes Alojadas en Internet */
	public function processImages($url, $extension) {

		/* Nombre Imagen */
		$nameImage = $this->nameFile . '_' . $this->i_img . '.' . $extension;

		/* Aumentar Valor del Iterador de la Imagen */
		$this->i_img++;

		/* Creacion del Archivo para Mover Datos de la Web a Local */
		$fp = fopen($this->directoryImages . $nameImage, "w");
		$ch = curl_init();
		curl_setopt($ch, CURLOPT_URL, $url);
		curl_setopt($ch, CURLOPT_HEADER, 0);
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
		curl_setopt($ch, CURLOPT_FILE, $fp);
		curl_exec($ch);
		curl_close($ch);
		fclose($fp);

		return (object) [
			'nameImage' => $nameImage, 
			'routeImage' => $this->directoryImages . $nameImage, 
		];
	}

	/* Parsiar el HTML */
	public function parseHTML(string $html) {

		/* Aplicar Remplazos en el HTML de Existir */
		if(count($this->replacements) > 0){
			foreach ($this->replacements as $key => $value) {
				$html = str_replace($key, $value, $html);
			}
		}

		/* Parserar HTML */
		$imagesData = '';
        $imagesNames = '';
		preg_match_all('/<img\s*(.*?)\s*src\s*=\s*"(.+?)"(.*?)>/u', $html, $matches);
		$i = 0;
		foreach($matches[0] as $imgTag) {
			if (str_contains($matches[2][$i], 'http') && $this->originImages == 'url') {
				/* DESCARGAR IMAGENES DE LA WEB */

				/* Extension Archivo */
				$imageExt = pathinfo($matches[2][$i])['extension'];

				/* Descargar Imagen y Alojarla Localmente en la Ruta del Sistema */
				$image = $this->processImages($matches[2][$i], $imageExt);
				$imageName = $image->nameImage;
				
				/* Ajustar Contenido Documento / Creando el XML con la Base64*/
				$imagesNames .=  '<o:File HRef=3D"'.$imageName.'"/>';
				$imageData = chunk_split(base64_encode(file_get_contents($image->routeImage)));

				/* Ajustar Datos en Documento */
				$imagesData .= '
					--'.$this->doc_file_separator.'
					Content-Location: images/'.$imageName.'
					Content-Transfer-Encoding: base64
					Content-Type: image/'.$imageExt.'

					'.$imageData.'
					';
				$imageDesc = '
					<v:imagedata src="images/'.$imageName.'" o:href=""/>
					</v:shape><![endif]--><![if !vml]><span style="mso-ignore:vglayout"><img border=3D0 src="images/'.$imageName.'"
					alt=3DHaut v:shapes="_x0000_i1057" /images/'.$imageName.' '.$matches[3][$i].'></span><![endif]>';

				$html = mb_ereg_replace($imgTag, $imageDesc, $html);

			} else if($this->originImages != 'url') {
				/* IMAGENES LOCALES */
				
				/* Extraer Data de Archivo Local */
				$dataImageLocal = pathinfo($matches[2][$i]);

				/* Extension Archivo */
				$imageExt = $dataImageLocal['extension'];

				/* Descargar Imagen y Alojarla Localmente en la Ruta del Sistema */
				$imageName = $dataImageLocal['basename'];
				
				/* Ajustar Contenido Documento / Creando el XML con la Base64*/
				$imagesNames .=  '<o:File HRef=3D"'.$imageName.'"/>';
				$imageData = chunk_split(base64_encode(file_get_contents($this->directoryImages . $imageName)));

				/* Ajustar Datos en Documento */
				$imagesData .= '
					--'.$this->doc_file_separator.'
					Content-Location: images/'.$imageName.'
					Content-Transfer-Encoding: base64
					Content-Type: image/'.$imageExt.'

					'.$imageData.'
					';
				$imageDesc = '
					<v:imagedata src="images/'.$imageName.'" o:href=""/>
					</v:shape><![endif]--><![if !vml]><span style="mso-ignore:vglayout"><img border=3D0 src="images/'.$imageName.'"
					alt=3DHaut v:shapes="_x0000_i1057" /images/'.$imageName.' '.$matches[3][$i].'></span><![endif]>';

				$html = mb_ereg_replace($imgTag, $imageDesc, $html);

			}

			/* Aumentar Iterador */
			$i++;
		}

		/* Remplazo de Caracteres en el HTML */
		$html = preg_replace('/=/u', '=3D', $html);

		/* Guardar Informacion del HTML */
		$this->html = [
				'html' => $html, 
				'imagesNames' => $imagesNames, 
				'imagesData' => $imagesData
			];
	}

	/* Creacion completa del Header Del Documento */
	public function headerDocument() {
		$head = 'MIME-Version: 1.0
			Content-Type: multipart/related; boundary="' . $this->doc_file_separator . '"

			--' . $this->doc_file_separator . '
			Content-Transfer-Encoding: quoted-printable
			Content-Type: text/html; charset="UTF-8"

			<html xmlns:o=3D"urn:schemas-microsoft-com:office:office"
			xmlns:w=3D"urn:schemas-microsoft-com:office:word"
			xmlns=3D"http://www.w3.org/TR/REC-html40">
			<head>
			<meta http-equiv=3DContent-Type content=3D"text/html; charset=3DUTF-8">
			<meta name=3DProgId content=3DWord.Document>
			<meta name=3DGenerator content=3D"Microsoft Word 11">
			<meta name=3DOriginator content=3D"Microsoft Word 11">
			<link rel=3DFile-List href=3D"filelist.xml">
			<!--[if gte mso 9]><xml>
				<w:WordDocument>
				<w:View>Print</w:View>
				<w:GrammarState>Clean</w:GrammarState>
				<w:ValidateAgainstSchemas/>
				<w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
				<w:IgnoreMixedContent>false</w:IgnoreMixedContent>
				<w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
				<w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
				</w:WordDocument>
			</xml><![endif]--><!--[if gte mso 9]><xml>
				<w:LatentStyles DefLockedState=3D"false" LatentStyleCount=3D"156">
				</w:LatentStyles>
			</xml><![endif]-->
			<style>
			<!--
				/* Style Definitions */
				p.MsoNormal, li.MsoNormal, div.MsoNormal
				{mso-style-parent:"";
				margin:0cm;
				margin-bottom:.0001pt;
				mso-pagination:widow-orphan;
				font-size:12.0pt;
				font-family:"Tahoma";
				mso-fareast-font-family:"Tahoma";}
			@page Section1
				{size:595.3pt 841.9pt;
				margin:1.5cm 1.5cm 1.5cm 1.5cm;
				mso-header-margin:35.4pt;
				mso-footer-margin:35.4pt;
				mso-paper-source:0;}
			div.Section1
				{page:Section1;}
			-->
			</style>
			<!--[if gte mso 10]>
			<style>
				/* Style Definitions */
				table.MsoNormalTable
				{mso-style-name:"\041E\0431\044B\0447\043D\0430\044F \0442\0430\0431\043B\=0438\0446\0430";
				mso-tstyle-rowband-size:0;
				mso-tstyle-colband-size:0;
				mso-style-noshow:yes;
				mso-style-parent:"";
				mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
				mso-para-margin:0cm;
				mso-para-margin-bottom:.0001pt;
				mso-pagination:widow-orphan;
				font-size:10.0pt;
				font-family:"Tahoma";
				mso-ansi-language:#0400;
				mso-fareast-language:#0400;
				mso-bidi-language:#0400;
				width:100%;
			}

			td.br1{
				border:1px solid black;
			}
			' . $this->stylesCSS . '
			</style>
			<![endif]-->
			</head>';
		return $head;
	}
	
	/* Creacion del Body */
	public function bodyDocument($body) {
		$body = '<body><div class=3D"Section1">'.$body.'</div>
			</body>
			</html>';
		return $body;
	}

	/* Lista de Archivos */
	public function fileList($entities) {
		$fileList = '
			--' . $this->doc_file_separator . '
			Content-Location: filelist.xml
			Content-Transfer-Encoding: quoted-printable
			Content-Type: text/xml; charset="utf-8"

			<xml xmlns:o=3D"urn:schemas-microsoft-com:office:office">
			<o:MainFile HRef=3D"../doc.doc"/>
			'.$entities['imagesNames'].'
			<o:File HRef=3D"filelist.xml"/>
			</xml>
			';
		return $fileList;
	}

	/* Footer Documento */
	public function footerDocument() {
		$footer = '</div>
			</body>
			</html>
			';
		return $footer;
	}

	/* Creacion del Documento */
	public function saveDoc() {

		/* Parciar HTML */
		$filePathEncoding = 'UTF-8';
		$entities = $this->html;
		$body = '';
		$body .= $this->headerDocument();
		$body .= $this->bodyDocument($entities['html']);
		$body .= $entities['imagesData'];
		$body .= $this->fileList($entities);
		$body .= '--' . $this->doc_file_separator . '--';
		$body = mb_ereg_replace('\\t+', '', $body);

		/* Ruta para Guardar el Archivo */
		$filePath = $this->fullPathFile;

		/* Definir donde se guardará el documento  */
		if($filePath) {
			if($filePathEncoding != $this->encoding){
				$filePath = iconv($this->encoding, $filePathEncoding, $filePath);
			} 
			file_put_contents($filePath, $body);
			return true;
		} else {
			return false;
		}

	}
	
	/* Retornar La Ruta del Documento */
	public function getRouteFile(){
		return $this->fullPathFile;
	}

}
