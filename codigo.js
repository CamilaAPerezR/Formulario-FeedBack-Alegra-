var MI_LLAVE_GEMINI = "AQ.Ab8RN6LG8KJTNNqRiXhS1bdEwIjyw0vjqE5xvcbEC6IoGP7j6A"; 

// Muestra el formulario web
function doGet() {
  var paginaWeb = HtmlService.createHtmlOutputFromFile('formulario');
  paginaWeb.setTitle('Captura de FeedBack - Alegra');
  return paginaWeb;
}

// Recibe los datos del formulario y los inserta en la hoja de cálculo
function guardarDatosEnSheets(nombre, producto, comentario) {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("FeedBack");
  var marcaTiempo = new Date();
  
  var usuario = nombre == "" ? "Anónimo" : nombre;
  
  // Insertamos la fila inicial con las columnas E y F vacías
  hoja.appendRow([marcaTiempo, producto, comentario, usuario, "", ""]);
  
  // Ejecutamos el análisis con la IA de Gemini
  analizarUltimaFilaConIA();
  
  // Al terminar con éxito, devolvemos el mensaje que activará el letrero verde
  return "Guardado exitoso";
}

var MI_LLAVE_GEMINI = "AQ.Ab8RN6LG8KJTNNqRiXhS1bdEwIjyw0vjqE5xvcbEC6IoGP7j6A"; 

// Muestra el formulario web
function doGet() {
  var paginaWeb = HtmlService.createHtmlOutputFromFile('formulario');
  paginaWeb.setTitle('Captura de FeedBack - Alegra');
  return paginaWeb;
}

// Recibe los datos del formulario y los inserta en la hoja de cálculo
function guardarDatosEnSheets(nombre, producto, comentario) {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("FeedBack");
  var marcaTiempo = new Date();
  
  var usuario = nombre == "" ? "Anónimo" : nombre;
  
  // Insertamos la fila inicial con las columnas E y F vacías
  hoja.appendRow([marcaTiempo, producto, comentario, usuario, "", ""]);
  
  // Ejecutamos el análisis con la IA de Gemini
  analizarUltimaFilaConIA();
  
  // Al terminar con éxito, devolvemos el mensaje que activará el letrero verde
  return "Guardado exitoso";
}

// FUNCIÓN DE IA: Llama a Gemini usando la API Key oficial del reto
function analizarUltimaFilaConIA() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("FeedBack");
  var ultimaFila = hoja.getLastRow();
  
  var comentarioUsuario = hoja.getRange(ultimaFila, 3).getValue();
  
  // 1. Endpoint oficial de Gemini con el parámetro ?key= al final
  var urlOficialGemini = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=" + MI_LLAVE_GEMINI;
  
  // Instrucciones súper claras para que la IA responda estructurado
  var instrucciones = "Analiza este comentario sobre el software Alegra: '" + comentarioUsuario + "'. " +
                      "Responde ÚNICAMENTE en este formato exacto: Sentimiento | Resumen " +
                      "Reglas: En la primera parte escribe solo una palabra (Positivo, Neutro o Negativo). " +
                      "En la segunda parte escribe un resumen corto de una frase. Sepáralos con una barra vertical '|'. " +
                      "No agregues asteriscos, saludos ni formato Markdown.";

  var cuerpoPeticion = {
    "contents": [{
      "parts": [{ "text": instrucciones }]
    }]
  };

  var opciones = {
    "method": "POST",
    "contentType": "application/json",
    "payload": JSON.stringify(cuerpoPeticion),
    "muteHttpExceptions": true
  };

   try {
    var respuestaServidor = UrlFetchApp.fetch(urlOficialGemini, opciones);
    var datosJson = JSON.parse(respuestaServidor.getContentText());
    
    // Validamos si Google devolvió un error en lugar de la respuesta
    if (datosJson.error) {
      hoja.getRange(ultimaFila, 5).setValue("Error API");
      hoja.getRange(ultimaFila, 6).setValue(datosJson.error.message);
      return;
    }
    
    // Si todo está bien, extrae el texto
    var textoIA = datosJson.candidates[0].content.parts[0].text;
    
    // Divide el texto usando tu estructura original de barra vertical
    var partes = textoIA.split("|");
    var sentimientoFinal = partes[0].trim();
    var resumenFinal = partes[1].trim();
    
    // Escribe los resultados reales en tus columnas E y F
    hoja.getRange(ultimaFila, 5).setValue(sentimientoFinal); 
    hoja.getRange(ultimaFila, 6).setValue(resumenFinal);     
    
  } catch(error) {
    // Si falla el código JavaScript en sí, esto te avisará
    hoja.getRange(ultimaFila, 5).setValue("Error JS");
    hoja.getRange(ultimaFila, 6).setValue(error.message);
  } 
}