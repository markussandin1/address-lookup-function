const {GoogleAuth} = require('google-auth-library');
const {google} = require('googleapis');
const axios = require('axios');
const htmlTemplates = require('./html-templates');
const sheetUtils = require('../address-lookup-local/sheet-utils');
const mapUtils = require('./map-utils');

// Huvudfunktion för Cloud Function
exports.addressLookup = async (req, res) => {
  // CORS-hantering
  res.set('Access-Control-Allow-Origin', '*');
  res.set('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  
  if (req.method === 'OPTIONS') {
    res.set('Access-Control-Allow-Headers', 'Content-Type');
    res.status(204).send('');
    return;
  }
  
  // Hämta sheetId från query-parametern eller request body
  const sheetId = req.query.sheetId || (req.body && req.body.sheetId);
  
  // Batch-parametrar (default till hela dokumentet om de inte anges)
  const startRow = parseInt(req.query.startRow || req.body?.startRow || 2); // Börja från rad 2 som default
  const batchSize = parseInt(req.query.batchSize || req.body?.batchSize || 100); // Bearbeta 100 rader åt gången som default
  
  // Ny parameter: skipFilled - hoppa över rader där adresskolumnen redan har ett värde (annat än "Hittade ingen adress")
  const skipFilled = req.query.skipFilled === 'true' || req.body?.skipFilled === true;
  
  // Kontrollera om klienten vill ha JSON (API-anrop) eller HTML (webbläsare)
  const wantsJson = req.headers.accept && 
                   (req.headers.accept.includes('application/json') || 
                    req.query.format === 'json');
  
  console.log('Request received. SheetId:', sheetId, 'startRow:', startRow, 
              'batchSize:', batchSize, 'skipFilled:', skipFilled, 'wantsJson:', wantsJson, 
              'UI mode:', req.query.ui === "1", 'Metadata request:', req.query.metadata === 'true',
              'Get addresses:', req.query.getAddresses === 'true');
  
  // Om inget sheetId angavs eller om detta är ett rent gränssnittsanrop
  if (!sheetId || (!wantsJson && req.query.ui === "1")) {
    return sendHtmlInterface(res, sheetId, startRow, batchSize, skipFilled);
  }
  
  // Om klienten vill hämta adresser för kartvisning
  if (req.query.getAddresses === 'true') {
    try {
      const addresses = await mapUtils.getAddressesWithCoordinates(sheetId);
      return res.status(200).json({ addresses });
    } catch (error) {
      console.error('Fel vid hämtning av adresser för karta:', error);
      return res.status(500).json({
        error: error.message || 'Ett fel uppstod vid hämtning av adresser för karta'
      });
    }
  }
  
  // Om klienten bara vill ha metadata om kalkylarket (för att visa UI snabbare)
  if (req.query.metadata === 'true') {
    try {
      const metadata = await sheetUtils.getSpreadsheetMetadata(sheetId);
      return res.status(200).json(metadata);
    } catch (error) {
      console.error('Fel vid hämtning av metadata:', error);
      return res.status(500).json({
        error: error.message || 'Ett fel uppstod vid hämtning av metadata'
      });
    }
  }
  
  try {
    // Skapa resultatobjektet
    const result = await sheetUtils.processSheet(sheetId, startRow, batchSize, skipFilled);
    
    // Om klienten vill ha JSON, skicka JSON
    if (wantsJson) {
      return res.status(200).json(result);
    }
    
    // Annars, skicka HTML-version av resultatet
    return sendHtmlResult(res, result, sheetId, skipFilled);
    
  } catch (error) {
    console.error('Oväntat fel:', error);
    
    const errorResult = {
      success: 0,
      failed: 0,
      empty: 0,
      skipped: 0,
      errors: [error.message || 'Ett oväntat fel inträffade'],
      batchInfo: {
        startRow: startRow,
        processedRows: 0,
        hasMoreRows: false,
        nextStartRow: startRow
      }
    };
    
    // Om klienten vill ha JSON, skicka JSON
    if (wantsJson) {
      return res.status(500).json(errorResult);
    }
    
    // Annars, skicka HTML-version av felet
    return sendHtmlError(res, error, sheetId, skipFilled);
  }
};

// Funktion för att skicka HTML-gränssnittet
function sendHtmlInterface(res, sheetId = "", startRow = 2, batchSize = 100, skipFilled = false) {
  console.log('Sending HTML interface. SheetId:', sheetId, 'skipFilled:', skipFilled);
  res.set('Content-Type', 'text/html');
  
  const safeSheetId = sheetId || '';
  const checkedAttribute = skipFilled ? 'checked' : '';
  
  // Använd HTML-mall från html-templates.js
  const html = htmlTemplates.getInterfaceTemplate(safeSheetId, checkedAttribute);
  
  res.status(200).send(html);
}

// Funktion för att skicka HTML-version av resultatet
function sendHtmlResult(res, result, sheetId, skipFilled) {
  console.log('Sending HTML result. SheetId:', sheetId, 'skipFilled:', skipFilled);
  res.set('Content-Type', 'text/html');
  
  const safeSheetId = sheetId || '';
  
  // Använd HTML-mall från html-templates.js
  const html = htmlTemplates.getResultTemplate(result, safeSheetId, skipFilled);
  
  res.status(200).send(html);
}

// Funktion för att skicka HTML-version av fel
function sendHtmlError(res, error, sheetId, skipFilled) {
  console.log('Sending HTML error. SheetId:', sheetId, 'skipFilled:', skipFilled);
  res.set('Content-Type', 'text/html');
  
  const safeSheetId = sheetId || '';
  
  // Använd HTML-mall från html-templates.js
  const html = htmlTemplates.getErrorTemplate(error, safeSheetId, skipFilled);
  
  res.status(500).send(html);
}