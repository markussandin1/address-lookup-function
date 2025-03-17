const {GoogleAuth} = require('google-auth-library');
const {google} = require('googleapis');
const axios = require('axios');

// Hjälpfunktion för att konvertera kolumnnummer till bokstav (t.ex. 0 -> A, 1 -> B)
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Funktion för att kontrollera och lägga till nödvändiga kolumner
async function setupSheetColumns(sheets, sheetId) {
  console.log('Kontrollerar om nödvändiga kolumner finns');
  
  try {
    // Hämta aktuella header-värden från rad 1
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: '1:1'  // Första raden där headers bör finnas
    });
    
    const headerRow = headerResponse.data.values ? headerResponse.data.values[0] : [];
    console.log('Befintliga kolumner: ', headerRow);
    
    // Kontrollera om de nödvändiga kolumnerna redan finns
    const hasAddressColumn = headerRow.includes('Adress');
    const hasLatitudeColumn = headerRow.includes('Latitud');
    const hasLongitudeColumn = headerRow.includes('Longitud');
    
    // Om vi inte har headers alls eller alla headers vi behöver finns redan
    if (headerRow.length === 0) {
      // Inga headers alls - lägg till alla
      console.log('Inga headers hittades, lägger till standardkolumner');
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: 'A1:D1',
        valueInputOption: 'RAW',
        resource: {
          values: [['Plats', 'Adress', 'Latitud', 'Longitud']]
        }
      });
      
      return {
        addressColumn: 'B',
        latColumn: 'C',
        lngColumn: 'D'
      };
    } 
    else {
      // Vi har headers - kontrollera vilka vi behöver lägga till
      const updates = [];
      let addressCol = headerRow.indexOf('Adress');
      let latCol = headerRow.indexOf('Latitud');
      let lngCol = headerRow.indexOf('Longitud');
      
      // Om någon kolumn saknas, lägg till den
      if (!hasAddressColumn) {
        // Hitta första tomma kolumnen eller lägg till i slutet
        const newAddressCol = headerRow.length;
        updates.push({
          range: `${columnToLetter(newAddressCol + 1)}1`,
          values: [['Adress']]
        });
        addressCol = newAddressCol;
      } else {
        addressCol = headerRow.indexOf('Adress');
      }
      
      if (!hasLatitudeColumn) {
        const newLatCol = addressCol + 1 + (hasAddressColumn ? 0 : 1);
        updates.push({
          range: `${columnToLetter(newLatCol + 1)}1`,
          values: [['Latitud']]
        });
        latCol = newLatCol;
      } else {
        latCol = headerRow.indexOf('Latitud');
      }
      
      if (!hasLongitudeColumn) {
        const newLngCol = latCol + 1 + (hasLatitudeColumn ? 0 : 1);
        updates.push({
          range: `${columnToLetter(newLngCol + 1)}1`,
          values: [['Longitud']]
        });
        lngCol = newLngCol;
      } else {
        lngCol = headerRow.indexOf('Longitud');
      }
      
      // Utför alla uppdateringar om det finns några
      if (updates.length > 0) {
        console.log('Lägger till saknade kolumner: ', updates);
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: sheetId,
          resource: {
            valueInputOption: 'RAW',
            data: updates
          }
        });
      }
      
      return {
        addressColumn: columnToLetter(addressCol + 1),
        latColumn: columnToLetter(latCol + 1),
        lngColumn: columnToLetter(lngCol + 1)
      };
    }
  } catch (error) {
    console.error('Fel vid kontroll av kolumner:', error);
    // Fallback till standardvärden om det blir fel
    return {
      addressColumn: 'B',
      latColumn: 'C',
      lngColumn: 'D'
    };
  }
}

// Funktion för att hämta metadata om kalkylarket
async function getSpreadsheetMetadata(sheetId) {
  // Skapa en autentiserad klient
  const auth = new GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  
  const client = await auth.getClient();
  const sheets = google.sheets({version: 'v4', auth: client});
  
  // Konfigurera kolumner först
  const columns = await setupSheetColumns(sheets, sheetId);
  
  // Hämta metadata från kalkylarket, nu med alla relevanta kolumner
  const range = `A:${columns.lngColumn}`;
  const metadataResponse = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    ranges: [range],
    fields: 'sheets.properties,sheets.data.rowData.values'
  });
  
  const totalRows = metadataResponse.data.sheets[0].properties.gridProperties.rowCount;
  const title = metadataResponse.data.sheets[0].properties.title;
  
  // Räkna antalet rader med data
  let rowsWithData = 0;
  let emptyRows = 0;
  let filledRows = 0;
  
  // Om data finns tillgänglig
  if (metadataResponse.data.sheets[0].data && 
      metadataResponse.data.sheets[0].data[0].rowData) {
    
    const rowData = metadataResponse.data.sheets[0].data[0].rowData;
    
    // Beräkna kolumnindex för adresskolumnen
    const addressColIndex = columns.addressColumn.charCodeAt(0) - 65;
    
    // Börja från rad 2 (index 1), hoppa över header-raden
    for (let i = 1; i < rowData.length; i++) {
      const row = rowData[i];
      
      // Om raden har värden
      if (row && row.values) {
        // Kontrollera om cell A har ett värde
        const hasValueA = row.values[0] && row.values[0].formattedValue;
        
        // Kontrollera om adresskolumnen har ett värde som inte är felmeddelande
        const hasValidAddress = row.values.length > addressColIndex && 
                              row.values[addressColIndex] && 
                              row.values[addressColIndex].formattedValue &&
                              row.values[addressColIndex].formattedValue !== 'Hittade ingen adress' &&
                              !row.values[addressColIndex].formattedValue.startsWith('Fel:');
        
        if (hasValueA) {
          rowsWithData++;
          
          if (hasValidAddress) {
            filledRows++;
          }
        } else {
          emptyRows++;
        }
      }
    }
  }
  
  return {
    title,
    totalRows,
    dataRows: rowsWithData,
    emptyRows,
    filledRows,
    needsUpdate: rowsWithData - filledRows,
    columns: columns
  };
}

// Funktion för att bearbeta en batch av kalkylarket
async function processSheet(sheetId, startRow, batchSize, skipFilled = false) {
  console.log('Bearbetar batch från rad ' + startRow + ', batchstorlek: ' + batchSize + ', skipFilled: ' + skipFilled);
  
  // Resultatstatistik
  const result = {
    success: 0,
    failed: 0,
    empty: 0,
    skipped: 0,
    errors: [],
    batchInfo: {
      startRow: startRow,
      processedRows: 0,
      hasMoreRows: false,
      nextStartRow: 0
    }
  };
  
  // Skapa en autentiserad klient
  const auth = new GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  
  const client = await auth.getClient();
  const sheets = google.sheets({version: 'v4', auth: client});
  
  // Kontrollera och konfigurera kolumner
  const columns = await setupSheetColumns(sheets, sheetId);
  console.log('Använder kolumner:', columns);
  
  // Hämta metadata från kalkylarket
  const metadataResponse = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    ranges: ['A:A'],  // Endast kolumn A för att avgöra storleken
    fields: 'sheets.properties'
  });
  
  const totalRows = metadataResponse.data.sheets[0].properties.gridProperties.rowCount;
  console.log('Kalkylarket har totalt ' + totalRows + ' rader');
  
  // Beräkna slutrad för denna batch
  const endRow = Math.min(startRow + batchSize - 1, totalRows);
  
  // Inkludera alla relevanta kolumner i hämtningen
  const range = `A${startRow}:${columns.lngColumn}${endRow}`;
  console.log('Hämtar data från range: ' + range);
  
  // Hämta data för aktuellt intervall
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: range
  });
  
  const rows = response.data.values || [];
  console.log('Hittade ' + rows.length + ' rader i denna batch');
  
  if (rows.length === 0) {
    result.errors.push('Inga data att bearbeta i detta intervall');
    result.batchInfo.hasMoreRows = false;
    result.batchInfo.nextStartRow = startRow;
    return result;
  }
  
  // Uppdateringsförfrågningar för batch-uppdatering
  const requests = [];
  
  // Google Maps API-nyckel
  const apiKey = process.env.MAPS_API_KEY || 'YOUR_API_KEY_HERE';
  
  // Beräkna kolumnindexer (0-baserat)
  const addressColIndex = columns.addressColumn.charCodeAt(0) - 65;
  
  // Loopa igenom alla rader i denna batch
  for (let i = 0; i < rows.length; i++) {
    const actualRow = startRow + i; // Den faktiska raden i kalkylarket
    const row = rows[i] || [];
    const searchTerm = row[0];
    
    // Befintliga värden (kan vara odefinierade om kolumnerna nyss har skapats)
    const existingAddress = row.length > addressColIndex ? row[addressColIndex] : undefined;
    
    // Om det inte finns något i kolumn A, hoppa över
    if (!searchTerm || searchTerm.toString().trim() === '') {
      console.log('Rad ' + actualRow + ': Tom, hoppar över');
      result.empty++;
      continue;
    }
    
    // Om skipFilled är true och det redan finns en adress som inte är "Hittade ingen adress", hoppa över
    if (skipFilled && 
        existingAddress && 
        existingAddress.toString().trim() !== '' && 
        existingAddress.toString().trim() !== 'Hittade ingen adress') {
      console.log('Rad ' + actualRow + ': Har redan adress, hoppar över');
      result.skipped++;
      continue;
    }
    
    try {
      // Använd Google Maps API för att hitta adressen
      console.log('Rad ' + actualRow + ': Söker adress för "' + searchTerm + '"');
      const geocodeUrl = 'https://maps.googleapis.com/maps/api/geocode/json?address=' + encodeURIComponent(searchTerm) + '&key=' + apiKey;
      const geocodeResponse = await axios.get(geocodeUrl);
      const geocodeData = geocodeResponse.data;
      
      let address = 'Hittade ingen adress';
      let latitude = '';
      let longitude = '';
      
      if (geocodeData.results && geocodeData.results.length > 0) {
        const geocodeResult = geocodeData.results[0];
        address = geocodeResult.formatted_address;
        
        // Hämta koordinater
        if (geocodeResult.geometry && geocodeResult.geometry.location) {
          latitude = geocodeResult.geometry.location.lat;
          longitude = geocodeResult.geometry.location.lng;
        }
        
        console.log('Rad ' + actualRow + ': Hittade adress "' + address + '" med koordinater Lat: ' + latitude + ', Lng: ' + longitude);
        result.success++;
      } else {
        console.log('Rad ' + actualRow + ': Hittade ingen adress');
        result.failed++;
      }
      
      // Lägg till värden i batch-uppdateringen med rätt kolumnnamn
      requests.push({
        range: `${columns.addressColumn}${actualRow}`,
        values: [[address]]
      });
      
      // Lägg till koordinater om vi hittade adressen
      if (latitude && longitude) {
        requests.push({
          range: `${columns.latColumn}${actualRow}:${columns.lngColumn}${actualRow}`,
          values: [[latitude, longitude]]
        });
      }
      
    } catch (error) {
      console.error('Fel vid sökning av "' + searchTerm + '" på rad ' + actualRow + ':', error.message);
      result.errors.push('Rad ' + actualRow + ': ' + (error.message || 'Okänt fel'));
      result.failed++;
      
      // Lägg till felmeddelande i batch-uppdatering med rätt kolumnnamn
      requests.push({
        range: `${columns.addressColumn}${actualRow}`,
        values: [['Fel: ' + (error.message || 'Okänt fel')]]
      });
    }
    
    // Lägg till en liten paus för att undvika API-begränsningar (men inte efter sista raden)
    if (i < rows.length - 1) {
      await new Promise(resolve => setTimeout(resolve, 200));
    }
  }
  
  // Uppdatera batch-info
  result.batchInfo.processedRows = rows.length;
  result.batchInfo.hasMoreRows = endRow < totalRows;
  result.batchInfo.nextStartRow = endRow + 1;
  
  // Utför batch-uppdateringen
  if (requests.length > 0) {
    console.log('Utför batch-uppdatering med ' + requests.length + ' förändringar');
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetId,
      resource: {
        valueInputOption: 'RAW',
        data: requests.map(req => ({
          range: req.range,
          values: req.values
        }))
      }
    });
  }
  
  console.log('Batch-bearbetning klar');
  return result;
}

module.exports = {
  columnToLetter,
  setupSheetColumns,
  getSpreadsheetMetadata,
  processSheet
};