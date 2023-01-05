<?php 

//Creacion Instancia Objeto. (Importar el USO de la Clase)
$msWord = new HtmlToWord(); 

//Tipo de Codificación (Opcional) por defecto UTF-8.
$msWord->encoding('UTF-8'); 

//Ruta Salida Del Archivo .DOC. (Garantizar que la carpeta exista antes de usar la libreria.)
$msWord->fileRoute(public_path() . '/documents'); 

//Nombre Archivo (Se Eliminaran Caracteres Especiales y Espacios por guión al piso.)
$msWord->nameFile('Documento'); 

// url : (Imagen alojaga en una URL que sea JPG o PNG) (Preferiblmente URL HTTPS para evitar errores)
// local : Imagenes Alojadas de Forma Local Unicamente Soporta PNG
$msWord->originImages('url'); 

// Ruta Carpeta Imagenes 
// (Obligatorio definirla, aca deben estar las imagenes locales, o si se manejarán por URL, en esta carpeta la libreria bajará las imagenes).
$msWord->routeImages(public_path(). '/images'); 

// Se eliminan los registros de antiguedad igual al segundo argumento en horas que esten dentro de los directorios definidos para imagenes y documentos 
// (Opcional, por defecto no borra el contenido de estos directorio).
$msWord->deleteOldFiles(true, 1); 

// Remplazos a ejecutar sobre el HTML original antes de parcearlo. 
// (Opcional (Primera columna es el valor a remplazar por el valor despues del =>))
$msWord->replacements([ 
    '<table border="0"' => '<table',
    '<br />' => '<br>',
    '<strong>' => '<b>',
    '</strong>' => '</b>',
]);

//Estilos CSS para aplicar al HTML (La llave es el elemento 'h1', la clase '.title' o el ID '#text')
$msWord->styles([ 
    'body' => 'font-size: 100%; font-family: "Cambria";',
    '.mi-clase' => 'width: 100%;',
    '#el_id' => 'width: 100%;',
]);

//HTML a procesar
$msWord->parseHTML($html); 

//Guardar El Documento En el Directorio Definido.
$msWord->saveDoc(); 

/* Obetener Ruta Del Archivo Word Para Descargarlo */
$msWord->getRouteFile();

?>