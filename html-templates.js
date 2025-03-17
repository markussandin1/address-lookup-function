// JavaScript för gränssnittet
const interfaceJsCode = `
  // Globala variabler
  let totalSuccess = 0;
  let totalFailed = 0;
  let totalEmpty = 0;
  let totalSkipped = 0;
  let totalRows = 0;
  let processedRows = 0;
  let allErrors = [];
  let batchesProcessed = 0;
  let sheetId = "SHEET_ID_PLACEHOLDER";
  let skipFilled = SKIP_FILLED_PLACEHOLDER;
  let metadata = null;
  let map = null;
  let mapMarkers = [];
  
  // Funktion för att extrahera sheet ID från en Google Sheets URL
  function extractSheetId(input) {
    try {
      if (!input) return "";
      
      // Om input redan är ett rent ID (inga slash, etc.) returnera som det är
      if (!/[\\/?=&#]/.test(input)) return input;
      
      const match = input.match(/\\/d\\/([^\\/?&#]+)/);
      if (match && match[1]) {
        console.log("Extracted sheet ID: " + match[1]);
        return match[1];
      }
      return input; // Returnera ursprunglig input om inget ID kunde hittas
    } catch (error) {
      console.error("Error extracting sheet ID:", error);
      return input;
    }
  }
  
  // När sidan laddas
  document.addEventListener("DOMContentLoaded", function() {
    // Lägg till eventlistener för ID-fältet för att hantera klistra in av URL:er
    document.getElementById("sheetId").addEventListener("input", function() {
      const input = this.value;
      if (input.indexOf("docs.google.com") !== -1) {
        const id = extractSheetId(input);
        if (id !== input) {
          this.value = id;
          console.log("URL detected and converted to sheet ID: " + id);
        }
      }
    });
    
    // Om URL:en innehåller parametrar
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.has("sheetId") && !urlParams.has("format") && !urlParams.has("ui")) {
      sheetId = urlParams.get("sheetId");
      document.getElementById("sheetId").value = sheetId;
      
      if (urlParams.has("skipFilled")) {
        skipFilled = urlParams.get("skipFilled") === "true";
        document.getElementById("skipFilled").checked = skipFilled;
      }
      
      // Visa loading-meddelande 
      document.getElementById("initialForm").style.display = "none";
      document.getElementById("loadingArea").style.display = "block";
      
      // Hämta metadata
      setTimeout(function() {
        fetchMetadata(sheetId);
      }, 100);
    }
    
    // Initiera Google Maps när den behövs
    if (document.getElementById('map')) {
      initMap();
    }
  });
  
  // Initiera kartan
  function initMap() {
    // Skapa en karta centrerad på Sverige
    map = new google.maps.Map(document.getElementById('map'), {
      center: { lat: 62.0, lng: 15.0 },
      zoom: 5
    });
  }
  
  // Visa alla adresser på kartan
  async function showAddressesOnMap() {
    if (!map) {
      document.getElementById('mapArea').style.display = 'block';
      initMap();
    }
    
    try {
      // Rensa befintliga markörer
      mapMarkers.forEach(marker => marker.setMap(null));
      mapMarkers = [];
      
      // Hämta adresser och koordinater från kalkylarket
      const metadataUrl = window.location.href.split("?")[0] + "?sheetId=" + sheetId + "&getAddresses=true";
      const response = await fetch(metadataUrl);
      
      if (!response.ok) {
        throw new Error("HTTP error! status: " + response.status);
      }
      
      const addressData = await response.json();
      console.log("Adressdata mottagen:", addressData);
      
      if (addressData.addresses && addressData.addresses.length > 0) {
        // Skapa bounds för att zooma in kartan
        const bounds = new google.maps.LatLngBounds();
        
        // Lägg till markörer för varje adress
        addressData.addresses.forEach(addr => {
          if (addr.lat && addr.lng) {
            const position = { lat: addr.lat, lng: addr.lng };
            const marker = new google.maps.Marker({
              position: position,
              map: map,
              title: addr.name || 'Plats'
            });
            
            // Lägg till infowindow med adressinfo
            const infoWindow = new google.maps.InfoWindow({
              content: '<div><strong>' + (addr.name || 'Plats') + '</strong><br>' + addr.address + '</div>'
            });
            
            marker.addListener('click', () => {
              infoWindow.open(map, marker);
            });
            
            mapMarkers.push(marker);
            bounds.extend(position);
          }
        });
        
        // Justera kartan för att visa alla markörer
        if (mapMarkers.length > 0) {
          map.fitBounds(bounds);
          // Om det bara finns en markör, zooma in mer
          if (mapMarkers.length === 1) {
            map.setZoom(14);
          }
        }
        
        document.getElementById('mapInfo').textContent = 'Visar ' + mapMarkers.length + ' platser på kartan';
      } else {
        document.getElementById('mapInfo').textContent = 'Inga platser med koordinater hittades i kalkylarket';
      }
    } catch (error) {
      console.error("Error loading addresses for map:", error);
      document.getElementById('mapInfo').textContent = 'Fel vid laddning av adresser: ' + error.message;
    }
  }
  
  // Formulärhantering
  document.getElementById("sheetForm").addEventListener("submit", function(e) {
    e.preventDefault();
    let inputVal = document.getElementById("sheetId").value.trim();
    
    // Kontrollera om det är en URL och extrahera ID om det är det
    sheetId = extractSheetId(inputVal);
    document.getElementById("sheetId").value = sheetId; // Uppdatera fältet med bara ID
    
    skipFilled = document.getElementById("skipFilled").checked;
    
    document.getElementById("initialForm").style.display = "none";
    document.getElementById("loadingArea").style.display = "block";
    
    fetchMetadata(sheetId);
  });
  
  // Hämta metadata
  async function fetchMetadata(sheetId) {
    try {
      console.log("Fetching metadata for sheetId: " + sheetId);
      const metadataUrl = window.location.href.split("?")[0] + "?sheetId=" + sheetId + "&metadata=true";
      console.log("Metadata URL: " + metadataUrl);
      const response = await fetch(metadataUrl);
      
      if (!response.ok) {
        throw new Error("HTTP error! status: " + response.status);
      }
      
      metadata = await response.json();
      console.log("Metadata received:", JSON.stringify(metadata));
      
      // Visa metadata
      document.getElementById("loadingArea").style.display = "none";
      document.getElementById("metadataArea").style.display = "block";
      document.getElementById("sheetTitle").textContent = metadata.title;
      document.getElementById("dataRows").textContent = metadata.dataRows;
      document.getElementById("needsUpdate").textContent = skipFilled ? metadata.needsUpdate : metadata.dataRows;
      document.getElementById("filledRows").textContent = metadata.filledRows;
      
      // Visa bekräfta-knapp
      document.getElementById("confirmButton").style.display = "block";
      document.getElementById("confirmButton").addEventListener("click", startProcessing);
      
      // Kontrollera om autoStart
      const urlParams = new URLSearchParams(window.location.search);
      if (urlParams.has("autoStart") && urlParams.get("autoStart") === "true") {
        console.log("Auto-starting processing");
        setTimeout(startProcessing, 1000);
      }
    } catch (error) {
      console.error("Error fetching metadata:", error);
      document.getElementById("loadingArea").style.display = "none";
      document.getElementById("initialForm").style.display = "block";
      alert("Ett fel uppstod: " + error.message + ". Försök igen.");
    }
  }
  
  // Starta bearbetningen
  function startProcessing() {
    console.log("Starting processing");
    // Dölj metadata och visa status
    document.getElementById("metadataArea").style.display = "none";
    document.getElementById("statusArea").style.display = "block";
    document.getElementById("results").style.display = "block";
    
    // Återställ räknare
    totalSuccess = 0;
    totalFailed = 0;
    totalEmpty = 0;
    totalSkipped = 0;
    processedRows = 0;
    allErrors = [];
    batchesProcessed = 0;
    
    // Uppskatta rader
    totalRows = skipFilled ? metadata.needsUpdate : metadata.dataRows;
    
    // Visa framstegsindikator
    document.getElementById("progressBarContainer").style.display = "block";
    
    // Starta första batchen
    processBatch(sheetId, 2, 100, skipFilled);
  }
  
  // Bearbeta batch
  async function processBatch(sheetId, startRow, batchSize, skipFilled) {
    console.log("Processing batch - startRow: " + startRow + ", batchSize: " + batchSize + ", skipFilled: " + skipFilled);
    batchesProcessed++;
    const batchNumber = batchesProcessed;
    
    // Lägg till batch i listan
    const batchItem = document.createElement("div");
    batchItem.id = "batch-" + batchNumber;
    batchItem.innerHTML = "Batch " + batchNumber + ": Bearbetar rad " + startRow + "-" + (startRow + batchSize - 1) + "... <span style='color: #007bff;'>[Pågår]</span>";
    document.getElementById("batchList").appendChild(batchItem);
    
    // Uppdatera statusmeddelande
    document.getElementById("statusMessage").innerHTML = "Bearbetar batch " + batchNumber + " (rad " + startRow + "-" + (startRow + batchSize - 1) + ")... Detta kan ta en stund.";
    
    try {
      // Anropa API:et
      const apiUrl = window.location.href.split("?")[0] + "?sheetId=" + sheetId + "&startRow=" + startRow + "&batchSize=" + batchSize + "&skipFilled=" + skipFilled + "&format=json";
      console.log("API URL: " + apiUrl);
      const response = await fetch(apiUrl);
      
      if (!response.ok) {
        throw new Error("HTTP error! status: " + response.status);
      }
      
      const data = await response.json();
      console.log("Batch processing result:", JSON.stringify(data));
      
      // Uppdatera resultatet
      totalSuccess += data.success;
      totalFailed += data.failed;
      totalEmpty += data.empty;
      totalSkipped += (data.skipped || 0);
      processedRows += data.batchInfo.processedRows;
      
      if (data.errors && data.errors.length > 0) {
        allErrors = allErrors.concat(data.errors);
      }
      
      // Uppdatera resultatvisningen
      document.getElementById("successCount").textContent = totalSuccess;
      document.getElementById("failedCount").textContent = totalFailed;
      document.getElementById("emptyCount").textContent = totalEmpty;
      document.getElementById("skippedCount").textContent = totalSkipped;
      
      // Visa fel om det finns några
      if (allErrors.length > 0) {
        document.getElementById("errors").style.display = "block";
        const errorList = document.getElementById("errorList");
        errorList.innerHTML = "";
        allErrors.forEach(function(error) {
          const li = document.createElement("li");
          li.textContent = error;
          errorList.appendChild(li);
        });
      }
      
      // Uppdatera batch-status
      document.getElementById("batch-" + batchNumber).innerHTML = 
        "Batch " + batchNumber + ": Rad " + startRow + "-" + (startRow + data.batchInfo.processedRows - 1) + " klart! ✅ (" + data.success + " hittade, " + data.failed + " misslyckade, " + (data.skipped || 0) + " behållna)";
      
      // Uppdatera framstegsindikatorn
      if (totalRows > 0) {
        const percentComplete = Math.min(100, Math.round(((totalSuccess + totalFailed + totalSkipped) / totalRows) * 100));
        document.getElementById("progress-bar").style.width = percentComplete + "%";
        document.getElementById("progressText").textContent = percentComplete + "% klart (" + (totalSuccess + totalFailed + totalSkipped) + " av uppskattade " + totalRows + " rader)";
      }
      
      // Om det finns fler rader
      if (data.batchInfo && data.batchInfo.hasMoreRows) {
        console.log("More rows to process. Next start row: " + data.batchInfo.nextStartRow);
        setTimeout(function() {
          processBatch(sheetId, data.batchInfo.nextStartRow, batchSize, skipFilled);
        }, 500);
      } else {
        console.log("Processing complete!");
        document.getElementById("statusMessage").className = "success";
        document.getElementById("statusMessage").innerHTML = 
          "Adressuppdateringen är klar! " + (totalSuccess + totalFailed + totalSkipped) + " rader har bearbetats. Du kan nu gå tillbaka till ditt kalkylark.";
        
        // Uppdatera till 100%
        document.getElementById("progress-bar").style.width = "100%";
        document.getElementById("progressText").textContent = "100% klart";
        
        // Lägg till länk
        const sheetLink = document.createElement("p");
        sheetLink.innerHTML = "<a href='https://docs.google.com/spreadsheets/d/" + sheetId + "/edit' target='_blank'>Öppna kalkylarket</a>";
        document.getElementById("statusMessage").appendChild(sheetLink);
        
        // Visa kartknappen nu när allt är klart
        document.getElementById('mapArea').style.display = 'block';
        document.getElementById('showMapButton').style.display = 'block';
      }
    } catch (error) {
      console.error("Error processing batch:", error);
      document.getElementById("batch-" + batchNumber).innerHTML = 
        "Batch " + batchNumber + ": Rad " + startRow + "-... misslyckades! ❌ (" + error.message + ")";
        
      document.getElementById("statusMessage").className = "error";
      document.getElementById("statusMessage").innerHTML = "Ett fel uppstod: " + error.message;
    }
  }
`;

// JavaScript för resultat-sidan
const resultJsCode = `
  // Globala variabler
  let sheetId = "SHEET_ID_PLACEHOLDER";
  let skipFilled = SKIP_FILLED_PLACEHOLDER;
  let map = null;
  let mapMarkers = [];
  
  // Initiera kartan
  function initMap() {
    // Skapa en karta centrerad på Sverige
    map = new google.maps.Map(document.getElementById('map'), {
      center: { lat: 62.0, lng: 15.0 },
      zoom: 5
    });
    
    // Om vi har ett giltigt kalkylark-ID, visa adresserna på kartan
    if (sheetId) {
      setTimeout(showAddressesOnMap, 1000);
    }
  }
  
  // Visa alla adresser på kartan
  async function showAddressesOnMap() {
    document.getElementById('loadingMap').style.display = 'block';
    document.getElementById('mapInfo').textContent = 'Laddar adresser...';
    
    try {
      // Rensa befintliga markörer
      mapMarkers.forEach(marker => marker.setMap(null));
      mapMarkers = [];
      
      // Hämta adresser och koordinater från kalkylarket
      const metadataUrl = window.location.href.split("?")[0] + "?sheetId=" + sheetId + "&getAddresses=true";
      const response = await fetch(metadataUrl);
      
      if (!response.ok) {
        throw new Error("HTTP error! status: " + response.status);
      }
      
      const addressData = await response.json();
      console.log("Adressdata mottagen:", addressData);
      
      document.getElementById('loadingMap').style.display = 'none';
      
      if (addressData.addresses && addressData.addresses.length > 0) {
        // Skapa bounds för att zooma in kartan
        const bounds = new google.maps.LatLngBounds();
        
        // Lägg till markörer för varje adress
        addressData.addresses.forEach(addr => {
          if (addr.lat && addr.lng) {
            const position = { lat: addr.lat, lng: addr.lng };
            const marker = new google.maps.Marker({
              position: position,
              map: map,
              title: addr.name || 'Plats'
            });
            
            // Lägg till infowindow med adressinfo
            const infoWindow = new google.maps.InfoWindow({
              content: '<div><strong>' + (addr.name || 'Plats') + '</strong><br>' + addr.address + '</div>'
            });
            
            marker.addListener('click', () => {
              infoWindow.open(map, marker);
            });
            
            mapMarkers.push(marker);
            bounds.extend(position);
          }
        });
        
        // Justera kartan för att visa alla markörer
        if (mapMarkers.length > 0) {
          map.fitBounds(bounds);
          // Om det bara finns en markör, zooma in mer
          if (mapMarkers.length === 1) {
            map.setZoom(14);
          }
        }
        
        document.getElementById('mapInfo').textContent = 'Visar ' + mapMarkers.length + ' platser på kartan';
      } else {
        document.getElementById('mapInfo').textContent = 'Inga platser med koordinater hittades i kalkylarket';
      }
    } catch (error) {
      console.error("Error loading addresses for map:", error);
      document.getElementById('loadingMap').style.display = 'none';
      document.getElementById('mapInfo').textContent = 'Fel vid laddning av adresser: ' + error.message;
    }
  }
`;

// Standard CSS stilar
const commonStyles = `
  body { font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 20px; color: #333; }
  .container { max-width: 800px; margin: 0 auto; background: #f9f9f9; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
  h1 { color: #2c3e50; text-align: center; }
  input[type="text"] { width: 100%; padding: 10px; font-size: 16px; margin-bottom: 10px; border: 1px solid #ddd; border-radius: 4px; }
  button { background: #3498db; color: white; border: none; padding: 10px 20px; font-size: 16px; cursor: pointer; border-radius: 4px; }
  button:hover { background: #2980b9; }
  .checkbox-container { display: flex; align-items: center; margin-bottom: 15px; }
  input[type="checkbox"] { margin-right: 10px; transform: scale(1.5); }
  .success { background: #d4edda; color: #155724; padding: 15px; border-radius: 4px; margin-top: 20px; }
  .error { background: #f8d7da; color: #721c24; padding: 15px; border-radius: 4px; margin-top: 20px; }
  .progress { background: #e0f7fa; padding: 15px; border-radius: 4px; margin-top: 20px; }
  .welcome { background: #fff8e1; padding: 15px; border-radius: 4px; margin-top: 20px; }
  .status-area { margin-top: 30px; }
  .loading-spinner { display: inline-block; width: 20px; height: 20px; border: 3px solid rgba(0,0,0,0.3); border-radius: 50%; border-top-color: #3498db; animation: spin 1s ease-in-out infinite; margin-right: 10px; }
  @keyframes spin { to { transform: rotate(360deg); } }
  table { width: 100%; border-collapse: collapse; margin-top: 20px; }
  th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
  th { background-color: #f2f2f2; }
  #progress-bar-container { height: 20px; background-color: #f0f0f0; border-radius: 10px; margin: 20px 0; overflow: hidden; }
  #progress-bar { height: 20px; background-color: #4CAF50; width: 0%; border-radius: 10px; transition: width 0.5s; }
  .metadata-area { background: #f0f8ff; padding: 15px; border-radius: 4px; margin-top: 20px; }
  .highlight { font-weight: bold; color: #3498db; }
  #map { height: 400px; width: 100%; border-radius: 4px; margin-top: 15px; }
  .map-area { background: #e8f5e9; padding: 15px; border-radius: 4px; margin-top: 20px; }
`;

// Funktion för att skapa HTML för gränssnittet
function getInterfaceTemplate(sheetId, checkedAttribute) {
  const jsCodeWithValues = interfaceJsCode
    .replace('SHEET_ID_PLACEHOLDER', sheetId)
    .replace('SKIP_FILLED_PLACEHOLDER', checkedAttribute ? 'true' : 'false');
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Adressuppdatering</title>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          ${commonStyles}
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Adressuppdatering för Google Kalkylark</h1>
          <p>Detta verktyg uppdaterar automatiskt adresser i ditt Google Kalkylark baserat på platsinformation i kolumn A. Koordinater sparas också för kartvisning.</p>
          
          <div id="initialForm">
            <form id="sheetForm">
              <label for="sheetId">Google Kalkylark ID eller URL:</label>
              <input type="text" id="sheetId" name="sheetId" value="${sheetId}" placeholder="t.ex. 1vYoG4kJ_FYupDBX16NxZljUblVZ3l_6xIlGFF1Ym0Nk eller hela URL:en" required>
              
              <div class="checkbox-container">
                <input type="checkbox" id="skipFilled" name="skipFilled" ${checkedAttribute}>
                <label for="skipFilled">Hoppa över rader där adress redan finns (behåll befintliga adresser)</label>
              </div>
              
              <button type="submit" id="startButton">Uppdatera adresser</button>
            </form>
          </div>
          
          <div id="loadingArea" class="status-area" style="display:none;">
            <div class="welcome">
              <div class="loading-spinner"></div>
              <span>Ansluter till kalkylarket och hämtar information...</span>
            </div>
          </div>
          
          <div id="metadataArea" class="metadata-area" style="display:none;">
            <h3>Kalkylarksinformation</h3>
            <p>Kalkylblad: <span id="sheetTitle" class="highlight">-</span></p>
            <p>Totalt antal rader med platsinformation: <span id="dataRows" class="highlight">-</span></p>
            <p>Rader som behöver uppdateras: <span id="needsUpdate" class="highlight">-</span></p>
            <p>Rader som redan har adresser: <span id="filledRows" class="highlight">-</span></p>
            
            <button id="confirmButton" style="display:none;">Starta uppdateringen</button>
          </div>
          
          <div id="statusArea" class="status-area" style="display:none;">
            <div id="statusMessage" class="progress"></div>
            
            <div id="progressBarContainer" style="display:none;">
              <h3>Total framgång:</h3>
              <div id="progress-bar-container">
                <div id="progress-bar"></div>
              </div>
              <p id="progressText">0% klart</p>
            </div>
            
            <div id="batchProgress">
              <h3>Batch-framsteg:</h3>
              <div id="batchList"></div>
            </div>
            
            <div id="results" style="display:none;">
              <h3>Resultat hittills:</h3>
              <table id="resultsTable">
                <tr>
                  <th>Typ</th>
                  <th>Antal</th>
                </tr>
                <tr>
                  <td>Framgångsrikt uppdaterade adresser</td>
                  <td id="successCount">0</td>
                </tr>
                <tr>
                  <td>Adresser som inte kunde hittas</td>
                  <td id="failedCount">0</td>
                </tr>
                <tr>
                  <td>Tomma rader som hoppades över</td>
                  <td id="emptyCount">0</td>
                </tr>
                <tr>
                  <td>Rader med befintliga adresser som behölls</td>
                  <td id="skippedCount">0</td>
                </tr>
              </table>
              
              <div id="errors" style="display:none;">
                <h3>Fel:</h3>
                <ul id="errorList"></ul>
              </div>
            </div>
          </div>
          
          <div id="mapArea" class="map-area" style="display:none;">
            <h3>Karta över platser</h3>
            <p id="mapInfo">Klicka för att visa dina adresser på kartan</p>
            <button id="showMapButton" onclick="showAddressesOnMap()">Visa adresser på karta</button>
            <div id="map"></div>
          </div>
        </div>
        
        <!-- Lägg till Google Maps JavaScript API -->
        <script src="https://maps.googleapis.com/maps/api/js?key=AIzaSyDebCtBpP2DI2V5cWnK6WuDmXu8seoVbF8&callback=initMap" async defer></script>
        
        <script>
          ${jsCodeWithValues}
        </script>
      </body>
    </html>
  `;
}

// Funktion för att skapa HTML för resultat
function getResultTemplate(result, sheetId, skipFilled) {
  const jsCodeWithValues = resultJsCode
    .replace('SHEET_ID_PLACEHOLDER', sheetId)
    .replace('SKIP_FILLED_PLACEHOLDER', skipFilled ? 'true' : 'false');
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Adressuppdatering - Resultat</title>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          ${commonStyles}
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Adressuppdatering för Google Kalkylark</h1>
          
          <div class="success">
            <h3>Uppdateringen är klar!</h3>
            <p>Adresser i ditt kalkylark har uppdaterats. Här är resultatet:</p>
          </div>
          
          <table>
            <tr>
              <th>Typ</th>
              <th>Antal</th>
            </tr>
            <tr>
              <td>Framgångsrikt uppdaterade adresser</td>
              <td>${result.success}</td>
            </tr>
            <tr>
              <td>Adresser som inte kunde hittas</td>
              <td>${result.failed}</td>
            </tr>
            <tr>
              <td>Tomma rader som hoppades över</td>
              <td>${result.empty}</td>
            </tr>
            <tr>
              <td>Rader med befintliga adresser som behölls</td>
              <td>${result.skipped || 0}</td>
            </tr>
          </table>
          
          <p style="margin-top: 20px;">
            <a href="https://docs.google.com/spreadsheets/d/${sheetId}/edit" target="_blank">Öppna kalkylarket</a> |
            <a href="?sheetId=${sheetId}&ui=1&skipFilled=${skipFilled}">Kör verktyget igen</a>
          </p>
          
          ${result.errors && result.errors.length > 0 ? `
          <div class="error">
            <h3>Fel under uppdateringen:</h3>
            <ul>
              ${result.errors.map(err => `<li>${err}</li>`).join('')}
            </ul>
          </div>
          ` : ''}
          
          <div class="map-area">
            <h3>Karta över platser</h3>
            <p id="mapInfo">Laddar karta...</p>
            <div id="loadingMap" style="margin-bottom: 10px;">
              <div class="loading-spinner"></div> Laddar adresser...
            </div>
            <div id="map"></div>
          </div>
        </div>
        
        <!-- Lägg till Google Maps JavaScript API -->
        <script src="https://maps.googleapis.com/maps/api/js?key=AIzaSyDebCtBpP2DI2V5cWnK6WuDmXu8seoVbF8&callback=initMap" async defer></script>
        
        <script>
          ${jsCodeWithValues}
        </script>
      </body>
    </html>
  `;
}

// Funktion för att skapa HTML för felmeddelanden
function getErrorTemplate(error, sheetId, skipFilled) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Adressuppdatering - Fel</title>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          ${commonStyles}
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Adressuppdatering för Google Kalkylark</h1>
          
          <div class="error">
            <h3>Ett fel uppstod</h3>
            <p>${error.message || 'Ett oväntat fel inträffade'}</p>
          </div>
          
          <p style="margin-top: 20px;">
            <a href="?sheetId=${sheetId}&ui=1&skipFilled=${skipFilled}">Försök igen</a>
          </p>
        </div>
      </body>
    </html>
  `;
}

module.exports = {
  getInterfaceTemplate,
  getResultTemplate,
  getErrorTemplate
};