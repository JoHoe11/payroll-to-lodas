// Ensure PDF.js worker is available
if (typeof pdfjsLib === 'undefined') {
    // Load PDF.js dynamically if not available
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js';
    script.onload = function() {
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
    };
    document.head.appendChild(script);
} else {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
}

document.addEventListener('DOMContentLoaded', function() {
    const startButton = document.getElementById('startButton');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const statusLabel = document.getElementById('statusLabel');
    const resultContainer = document.getElementById('resultContainer');
    const resultMessage = document.getElementById('resultMessage');
    const downloadLink = document.getElementById('downloadLink');
    
    // Validate filename contains G2N
    const outputFilename = document.getElementById('outputFilename');
    outputFilename.addEventListener('change', function() {
        if (!this.value.includes('G2N')) {
            alert('Der Dateiname sollte die Kennung G2N enthalten.');
        }
    });

    //UI-User Input
    startButton.addEventListener('click', async function() {
        // Get form values
        const monat = document.getElementById('monat').value;
        const bAVspecialcode = document.getElementById('bAVspecialcode').value;
        const NBAspecialcode = document.getElementById('NBAspecialcode').value;
        const outputFilename = document.getElementById('outputFilename').value;
        
        // Get files
        const pdfFile = document.getElementById('pdfFile').files[0];
        const lohnjournalFile = document.getElementById('lohnjournalFile').files[0];
        const personalkostenFile = document.getElementById('personalkostenFile').files[0];
        
        // Validate inputs
        if (!pdfFile || !lohnjournalFile || !personalkostenFile) {
            alert('Bitte wählen Sie alle erforderlichen Dateien aus.');
            return;
        }
        
        // Show progress
        startButton.disabled = true;
        progressContainer.classList.remove('hidden');
        
        try {
            // 1. Process PDF on the client side
            statusLabel.textContent = 'PDF wird analysiert...';
            progressBar.style.width = '20%';
            const pdfData = await processPDF(pdfFile, monat, bAVspecialcode, NBAspecialcode);
            
            // 2. Process Excel files
            statusLabel.textContent = 'Excel-Dateien werden geladen...';
            progressBar.style.width = '40%';
            const lohnjournalData = await readExcelFile(lohnjournalFile);
            const personalkostenData = await readExcelFile(personalkostenFile);
            
            // 3. Merge data
            statusLabel.textContent = 'Daten werden zusammengeführt...';
            progressBar.style.width = '60%';
            const mergedData = mergeData(pdfData, lohnjournalData, personalkostenData, monat);
            
            // 4. Create Excel file
            statusLabel.textContent = 'Excel-Datei wird erstellt...';
            progressBar.style.width = '80%';
            const excelBlob = createExcelFile(mergedData, outputFilename);
            
            // 5. Generate download link directly
            progressBar.style.width = '100%';
            statusLabel.textContent = 'Verarbeitung wurde erfolgreich abgeschlossen.';
            resultMessage.textContent = 'Daten wurden erfolgreich verarbeitet. Klicken Sie auf den Link, um die Excel-Datei herunterzuladen.';
            
            // Create URL for download
            const url = URL.createObjectURL(excelBlob);
            downloadLink.href = url;
            downloadLink.download = outputFilename;
            resultContainer.classList.remove('hidden');
            
        } catch (error) {
            console.error('Error:', error);
            progressBar.style.width = '100%';
            statusLabel.textContent = 'Fehler bei der Verarbeitung.';
            alert('Ein Fehler ist aufgetreten: ' + error.message);
        } finally {
            startButton.disabled = false;
        }
    });
    
    async function processPDF(pdfFile, selectedMonth, bAVspecialcode, NBAspecialcode) {
        // Initialize data arrays
        const nba_data = [];
        const law_data = [];
        
        // Store tax data by personnel number
        const taxDataByPerson = {};
        
        // Use Sets to track unique combinations with month awareness
        const seenLawCombos = new Set();
        const seenNbaCombos = new Set();
        
        // Read PDF file
        const arrayBuffer = await pdfFile.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        console.log('PDF loaded. Pages:', pdf.numPages);
        console.log('Selected month:', selectedMonth);
        
        // Process each page
        for (let i = 1; i <= pdf.numPages; i++) {
            try {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                
                // Convert textContent to a structured format
                let textItems = textContent.items.map(item => ({
                    text: item.str,
                    x: item.transform[4],
                    y: item.transform[5]
                }));
                
                // Sort items by position on page for easy reading
                textItems.sort((a, b) => {
                    // Sort by y coordinate (top to bottom)
                    if (Math.abs(a.y - b.y) > 5) {
                        return b.y - a.y; // Descending order for y
                    }
                    // If y is similar, sort by x (left to right)
                    return a.x - b.x;
                });
                
                // Join the text for pattern matching
                const text = textItems.map(item => item.text).join(' ');
                
                // 1. Detect month on the page
                const pageMonth = detectMonth(text);
                const isSelectedMonth = monthsMatch(pageMonth, selectedMonth);
                console.log(`Page ${i} - Month: ${pageMonth || 'unknown'}, Matches selected month: ${isSelectedMonth}`);
                
                // 2. Look for personnel number
                const personalnummer = extractPersonalNumber(text);
                if (!personalnummer) {
                    console.log(`Page ${i} - No personnel number found, skipping page`);
                    continue;
                }
                
                console.log(`Page ${i} - Personnel number: ${personalnummer}`);
                
                // 3. Extract tax data using improved function
                const taxData = extractTaxDataByPosition(textItems, personalnummer, isSelectedMonth);
                
                // Store tax data for this person
                if (!taxDataByPerson[personalnummer]) {
                    taxDataByPerson[personalnummer] = {
                        currentTax: 0,
                        backdatedTax: 0,
                        backdatedInsurance: 0
                    };
                }
                
                // Only update current tax if this is the selected month
                if (isSelectedMonth) {
                    taxDataByPerson[personalnummer].currentTax = taxData.currentTax;
                }
                
                // Always update backdated values if they exist (they might appear on any page)
                if (taxData.backdatedTax > 0) {
                    taxDataByPerson[personalnummer].backdatedTax = taxData.backdatedTax;
                }
                
                if (taxData.backdatedInsurance > 0) {
                    taxDataByPerson[personalnummer].backdatedInsurance = taxData.backdatedInsurance;
                }
                
                // 4. Extract Lohnart data (LAW section)
                let lawSection = "";
                if (text.includes("Lohnart")) {
                    let startIdx = text.indexOf("Lohnart");
                    let endIdx = text.indexOf("Gesamt-Brutto");
                    
                    if (endIdx === -1) {
                        // Try to find a different ending marker
                        endIdx = text.indexOf("Gesamt");
                        if (endIdx === -1) {
                            // If no ending marker, take a chunk of text
                            endIdx = startIdx + 1000;
                        }
                    }
                    
                    if (startIdx < endIdx) {
                        lawSection = text.substring(startIdx, endIdx);
                        console.log(`Page ${i} - LAW section found, length: ${lawSection.length}`);
                    }
                }
                
                if (lawSection) {
                    if (isSelectedMonth) {
                        // Extract regular LAW entries for selected month
                        const regularEntries = extractRegularLawEntries(lawSection, personalnummer);
                        console.log(`Page ${i} - Found ${regularEntries.length} regular LAW entries`);
                        
                        // Add to law_data array, tracking by month-aware key
                        for (const entry of regularEntries) {
                            const key = `${entry.Personalnummer}-${entry.Lohnart}-${entry["Lohnart Betrag"]}-${pageMonth}`;
                            if (!seenLawCombos.has(key)) {
                                law_data.push(entry);
                                seenLawCombos.add(key);
                            } else {
                                console.log(`Skipping duplicate LAW entry: ${key}`);
                            }
                        }
                    }
                    
                    // Always extract N-prefixed LAW entries
                    const nEntries = extractNPrefixedLawEntries(lawSection, personalnummer);
                    console.log(`Page ${i} - Found ${nEntries.length} N-prefixed LAW entries`);
                    
                    // Add to law_data array, tracking by month-aware key
                    for (const entry of nEntries) {
                        const key = `${entry.Personalnummer}-${entry.Lohnart}-${entry["Lohnart Betrag"]}-${pageMonth}`;
                        if (!seenLawCombos.has(key)) {
                            law_data.push(entry);
                            seenLawCombos.add(key);
                        } else {
                            console.log(`Skipping duplicate N-prefixed LAW entry: ${key}`);
                        }
                    }
                }
                
                // 5. Extract NBA data
                let nbaSection = "";
                if (text.includes("Netto-Bezüge/Netto-Abzüge")) {
                    let startIdx = text.indexOf("Netto-Bezüge/Netto-Abzüge");
                    let endIdx = text.indexOf("Bank", startIdx);
                    
                    if (endIdx === -1) {
                        endIdx = text.indexOf("Verdienst", startIdx);
                        if (endIdx === -1) {
                            endIdx = startIdx + 1000;
                        }
                    }
                    
                    if (startIdx < endIdx) {
                        nbaSection = text.substring(startIdx, endIdx);
                        console.log(`Page ${i} - NBA section found, length: ${nbaSection.length}`);
                    }
                } else if (text.includes("Netto-Bezüge") || text.includes("Netto-Abzüge")) {
                    let startIdx = text.includes("Netto-Bezüge") ? 
                        text.indexOf("Netto-Bezüge") : text.indexOf("Netto-Abzüge");
                    
                    let endIdx = text.indexOf("Bank", startIdx);
                    if (endIdx === -1) {
                        endIdx = text.indexOf("Verdienst", startIdx);
                        if (endIdx === -1) {
                            endIdx = startIdx + 1000;
                        }
                    }
                    
                    if (startIdx < endIdx) {
                        nbaSection = text.substring(startIdx, endIdx);
                        console.log(`Page ${i} - NBA section found, length: ${nbaSection.length}`);
                    }
                }
                
                if (nbaSection) {
                    if (isSelectedMonth) {
                        // Extract regular NBA entries for selected month
                        const regularEntries = extractRegularNbaEntries(nbaSection, personalnummer);
                        console.log(`Page ${i} - Found ${regularEntries.length} regular NBA entries`);
                        
                        // Add to nba_data array, tracking by month-aware key
                        for (const entry of regularEntries) {
                            const key = `${entry.Personalnummer}-${entry["Netto-Bezüge/ -Abzüge Nummer"]}-${entry["Netto-Bezüge/ -Abzüge Betrag"]}-${pageMonth}`;
                            if (!seenNbaCombos.has(key)) {
                                nba_data.push(entry);
                                seenNbaCombos.add(key);
                            } else {
                                console.log(`Skipping duplicate NBA entry: ${key}`);
                            }
                        }
                    }
                    
                    // Always extract N-prefixed NBA entries
                    const nEntries = extractNPrefixedNbaEntries(nbaSection, personalnummer);
                    console.log(`Page ${i} - Found ${nEntries.length} N-prefixed NBA entries`);
                    
                    // Add to nba_data array, tracking by month-aware key
                    for (const entry of nEntries) {
                        const key = `${entry.Personalnummer}-${entry["Netto-Bezüge/ -Abzüge Nummer"]}-${entry["Netto-Bezüge/ -Abzüge Betrag"]}-${pageMonth}`;
                        if (!seenNbaCombos.has(key)) {
                            nba_data.push(entry);
                            seenNbaCombos.add(key);
                        } else {
                            console.log(`Skipping duplicate N-prefixed NBA entry: ${key}`);
                        }
                    }
                    
                    // Always extract "already paid" entries
                    const alreadyPaidEntries = extractAlreadyPaidEntries(nbaSection, personalnummer);
                    console.log(`Page ${i} - Found ${alreadyPaidEntries.length} "already paid" entries`);
                    
                    // Add to nba_data array, tracking by month-aware key
                    for (const entry of alreadyPaidEntries) {
                        const key = `${entry.Personalnummer}-${entry["Netto-Bezüge/ -Abzüge Nummer"]}-${entry["Netto-Bezüge/ -Abzüge Betrag"]}-${pageMonth}`;
                        if (!seenNbaCombos.has(key)) {
                            nba_data.push(entry);
                            seenNbaCombos.add(key);
                        } else {
                            console.log(`Skipping duplicate already paid entry: ${key}`);
                        }
                    }
                    
                    // Extract "aus NB" entries
                    const ausNbEntries = extractAusNbEntries(nbaSection, personalnummer);
                    console.log(`Page ${i} - Found ${ausNbEntries.length} "aus NB" entries`);
                    
                    // Add to nba_data array, tracking by month-aware key
                    for (const entry of ausNbEntries) {
                        const key = `${entry.Personalnummer}-${entry["Netto-Bezüge/ -Abzüge Nummer"]}-${entry["Netto-Bezüge/ -Abzüge Betrag"]}-${pageMonth}`;
                        if (!seenNbaCombos.has(key)) {
                            nba_data.push(entry);
                            seenNbaCombos.add(key);
                        } else {
                            console.log(`Skipping duplicate aus NB entry: ${key}`);
                        }
                    }
                }
                
            } catch (error) {
                console.error(`Error processing page ${i}:`, error);
            }
        }
        
        // Convert tax data to array for merging
        const tax_social_calculated = [];
        Object.entries(taxDataByPerson).forEach(([persNr, data]) => {
            tax_social_calculated.push({
                "Pers.Nr.": parseInt(persNr),
                "Steuer aktuell": data.currentTax,
                "Steuer backdated Summe": data.backdatedTax,
                "Versicherung backdated Summe": data.backdatedInsurance
            });
        });
        
        console.log('PDF processing complete:');
        console.log('- NBA data:', nba_data.length, 'entries');
        console.log('- LAW data:', law_data.length, 'entries');
        console.log('- Tax/Social data:', tax_social_calculated.length, 'entries');
        console.log('Tax data by person:', taxDataByPerson);
        
        return {
            nba_data,
            law_data,
            tax_social_calculated
        };
    }
    
    // Function to extract tax data with focus on far right column for backdated values
function extractTaxDataByPosition(textItems, personalnummer, isSelectedMonth) {
    console.log(`Extracting tax data for ${personalnummer} using improved column-based approach`);
    
    const result = {
        currentTax: 0,
        backdatedTax: 0,
        backdatedInsurance: 0
    };
    
    // Find key phrases and their positions
    let steuerHeaderItem = null;
    let svHeaderItem = null;
    let nMarkerItems = [];
    let rightColumnItems = [];
    
    // Get page dimensions by finding max coordinates
    const maxX = Math.max(...textItems.map(item => item.x));
    const minX = Math.min(...textItems.map(item => item.x));
    const pageWidth = maxX - minX;
    
    // Define the right column as items that are in the rightmost 20% of the page
    const rightColumnThreshold = maxX - (pageWidth * 0.20);
    
    // Log all text items for debugging
    console.log(`Found ${textItems.length} text items for ${personalnummer}`);
    console.log(`Page width: ${pageWidth}, right column threshold: ${rightColumnThreshold}`);
    
    // First pass: find the key headers and items
    for (let i = 0; i < textItems.length; i++) {
        const item = textItems[i];
        const text = item.text.toLowerCase();
        
        // Collect items in the far right column
        if (item.x > rightColumnThreshold) {
            rightColumnItems.push(item);
        }
        
        // For Steuerrechtliche Abzüge - try multiple variations
        if (text.includes("steuerrechtliche abzüge") || 
            text.includes("steuerrechtliche abzuge") ||
            text.includes("steuerrechtliche abz") ||
            text.includes("steuer abz")) {
            steuerHeaderItem = item;
            console.log(`Found steuerrechtliche header: "${item.text}" at (${item.x}, ${item.y})`);
        } 
        // For SV-rechtliche Abzüge - try multiple variations
        else if (text.includes("sv-rechtliche abzüge") || 
                text.includes("sv-rechtliche abzuge") ||
                text.includes("sv-rechtliche abz") ||
                text.includes("sv rechtliche") ||
                text.includes("sv-rechtl")) {
            svHeaderItem = item;
            console.log(`Found sv-rechtliche header: "${item.text}" at (${item.x}, ${item.y})`);
        }
        // Look for "N" markers which indicate backdated entries
        else if ((text === "n" || text === "n " || text.trim() === "n") && item.x < (pageWidth * 0.3)) {
            nMarkerItems.push(item);
            console.log(`Found N marker at (${item.x}, ${item.y})`);
        }
    }
    
    // Function to extract a number from text with proper decimal handling
    const extractNumber = (text) => {
        if (!text) return 0;
        
        const matches = text.match(/[\d.,]+/g);
        if (matches && matches.length > 0) {
            // Get the last match as it's likely the amount
            const lastMatch = matches[matches.length - 1];
            
            // Try to determine if this is a currency number by looking for patterns
            const isCurrency = /\d+,\d{2}$/.test(lastMatch) || // Ends with ,XX
                              /\d+\.\d{2}$/.test(lastMatch) || // Ends with .XX
                              lastMatch.includes('.') && lastMatch.includes(','); // Has both . and ,
            
            let value;
            if (isCurrency) {
                // This is likely a currency - handle German format (1.234,56)
                value = parseFloat(lastMatch.replace(/\./g, '').replace(',', '.'));
            } else {
                // This might be a normal number - try to interpret
                value = parseFloat(lastMatch.replace(',', '.'));
            }
            
            console.log(`Extracted number: ${value} from text: "${text}" (original: ${lastMatch})`);
            return value;
        }
        return 0;
    };
    
    // STEP 1: Find Steuer aktuell (current tax) - this part works correctly
    if (steuerHeaderItem) {
        // Find the actual tax value - look both in the right column AND under the header
        const possibleTaxItems = rightColumnItems.filter(item => 
            // Slightly below the header
            item.y < steuerHeaderItem.y &&
            // Not too far below
            (steuerHeaderItem.y - item.y) < 100 &&
            // Contains digits
            /\d/.test(item.text)
        );
        
        if (possibleTaxItems.length > 0) {
            // Sort by vertical proximity to header
            possibleTaxItems.sort((a, b) => (steuerHeaderItem.y - a.y) - (steuerHeaderItem.y - b.y));
            
            // The closest item should be our value
            result.currentTax = extractNumber(possibleTaxItems[0].text);
            console.log(`Found Steuer aktuell: ${result.currentTax} from "${possibleTaxItems[0].text}"`);
        }
    }
    
    // STEP 2: Find Steuer backdated and Versicherung backdated values
    if (nMarkerItems.length > 0) {
        // Sort N markers by vertical position (top to bottom)
        nMarkerItems.sort((a, b) => b.y - a.y);
        
        // Process each N marker
        nMarkerItems.forEach((nItem, index) => {
            // Find right column items that are at approximately the same vertical position as this N marker
            const alignedRightItems = rightColumnItems.filter(item => 
                Math.abs(item.y - nItem.y) < 15 && // Within 15 units vertically
                /\d/.test(item.text) // Contains digits
            );
            
            if (alignedRightItems.length > 0) {
                // Take the item with the most rightward position
                alignedRightItems.sort((a, b) => b.x - a.x);
                const value = extractNumber(alignedRightItems[0].text);
                
                // First N marker is for tax backdated
                if (index === 0) {
                    result.backdatedTax = value;
                    console.log(`Found Steuer backdated: ${result.backdatedTax} from "${alignedRightItems[0].text}"`);
                }
                // Last N marker is for insurance backdated (might be the same as first if only one N)
                else if (index === nMarkerItems.length - 1) {
                    result.backdatedInsurance = value;
                    console.log(`Found Versicherung backdated: ${result.backdatedInsurance} from "${alignedRightItems[0].text}"`);
                }
            }
        });
        
        // Special case: If we have only one N marker, check if there's a value at the bottom right
        // that could be the insurance backdated value
        if (nMarkerItems.length === 1 && result.backdatedTax > 0 && result.backdatedInsurance === 0) {
            const bottomRightItems = rightColumnItems.filter(item =>
                item.y < nMarkerItems[0].y - 20 && // Significantly below the N marker
                /\d/.test(item.text) // Contains digits
            );
            
            if (bottomRightItems.length > 0) {
                // Sort by vertical position (bottom-most first)
                bottomRightItems.sort((a, b) => a.y - b.y);
                result.backdatedInsurance = extractNumber(bottomRightItems[0].text);
                console.log(`Found likely Versicherung backdated at bottom: ${result.backdatedInsurance} from "${bottomRightItems[0].text}"`);
            }
        }
    }
    
    // Ensure all values are properly rounded to 2 decimal places
    result.currentTax = Math.round(result.currentTax * 100) / 100;
    result.backdatedTax = Math.round(result.backdatedTax * 100) / 100;
    result.backdatedInsurance = Math.round(result.backdatedInsurance * 100) / 100;
    
    console.log(`Final tax data for ${personalnummer}:`, result);
    return result;
}
    
    // Helper function to detect month from text
    function detectMonth(text) {
        // Common patterns for month extraction
        const patterns = [
            /für\s+(\w+)\s+\d{4}/i,                                // "für Juni 2025"
            /Abrechnung der Brutto\/Netto-Bezüge für (\w+)\s+\d{4}/i, // Full header pattern
            /Abrechnung.*für\s+(\w+)\s+\d{4}/i,                    // Generic "Abrechnung ... für Month Year"
            /Monat\s*:\s*(\w+)/i,                                  // "Monat: Juni"
            /Periode\s*:\s*(\w+)/i                                 // "Periode: Juni"
        ];
        
        for (const pattern of patterns) {
            const match = text.match(pattern);
            if (match && match[1]) {
                return match[1].trim();
            }
        }
        
        // If no pattern matches, try to find month names in the text
        const months = [
            "Januar", "Februar", "März", "April", "Mai", "Juni", 
            "Juli", "August", "September", "Oktober", "November", "Dezember"
        ];
        
        const firstParagraph = text.substring(0, 300).toLowerCase();
        for (const month of months) {
            if (firstParagraph.includes(month.toLowerCase())) {
                return month;
            }
        }
        
        return null;
    }
    
    // Helper function to match months
    function monthsMatch(month1, month2) {
        if (!month1 || !month2) return false;
        return month1.toLowerCase() === month2.toLowerCase();
    }
    
    // Helper function to extract personnel number
    function extractPersonalNumber(text) {
        const match = text.match(/\*Pers\.-Nr\.\s(\d{5})/);
        return match ? match[1] : null;
    }
    
    // Extract regular LAW entries (without N prefix)
    function extractRegularLawEntries(section, personalnummer) {
        const result = [];
        
        // Clean up whitespace
        section = section.replace(/\s+/g, ' ').trim();
        
        // Various patterns to try
        const patterns = [
            // Pattern 1: Standard format with codes
            /\s(\d{3,4})\s+(.+?)\s+([LJSAFEPVWMN])(?:\s+([LJSAFEPVWMN]))?(?:\s+([LJSAFEPVWMN]))?\s+([\d.,]+)(-?)/g,
            
            // Pattern 2: Without codes
            /\s(\d{3,4})\s+([^\d]+?)\s+([\d.,]+)(-?)/g,
            
            // Pattern 3: Very general pattern
            /\b(\d{3,4})\b\s+([^\d]+?)\s+\b([\d.,]+)\b/g
        ];
        
        for (const pattern of patterns) {
            let match;
            let matchCount = 0;
            
            // Reset pattern index
            pattern.lastIndex = 0;
            
            while ((match = pattern.exec(section)) !== null) {
                let nummer, bezeichnung, betrag, vorzeichen;
                
                if (pattern === patterns[0]) {
                    // Pattern 1 with codes
                    nummer = match[1];
                    bezeichnung = match[2].trim();
                    betrag = match[6].replace(/\./g, '').replace(',', '.');
                    vorzeichen = match[7] || '';
                } else if (pattern === patterns[1]) {
                    // Pattern 2 without codes
                    nummer = match[1];
                    bezeichnung = match[2].trim();
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = match[4] || '';
                } else {
                    // Pattern 3: Very general
                    nummer = match[1];
                    bezeichnung = match[2].trim();
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = '';
                }
                
                // Skip N prefixed entries
                if (nummer.startsWith('N') || nummer.startsWith('n')) continue;
                
                // Format the amount
                const betragValue = vorzeichen === '-' ? -parseFloat(betrag) : parseFloat(betrag);
                
                if (!isNaN(betragValue)) {
                    matchCount++;
                    result.push({
                        "Personalnummer": personalnummer,
                        "Lohnart": nummer,
                        "Lohnart Bezeichnung": bezeichnung,
                        "Lohnart Betrag": betragValue
                    });
                }
            }
            
            // If found matches with this pattern, no need to try others
            if (matchCount > 0) break;
        }
        
        return result;
    }
    
    // Extract N-prefixed LAW entries
    function extractNPrefixedLawEntries(lawSection, personalnummer) {
        const result = [];
        
        // Clean up whitespace
        lawSection = lawSection.replace(/\s+/g, ' ').trim();
        
        // Various patterns to try - with improved space handling between N and digits
        const patterns = [
            // Pattern 1: With N prefix and codes - allowing for variable spaces
            /N\s*(\d{3,4})\s+(.+?)\s+([LJSAFEPVWMN])(?:\s+([LJSAFEPVWMN]))?(?:\s+([LJSAFEPVWMN]))?\s+([\d.,]+)(-?)/g,
            
            // Pattern 2: With N prefix, without codes - allowing for variable spaces
            /N\s*(\d{3,4})\s+([^\d]+?)\s+([\d.,]+)(-?)/g,
            
            // Pattern 3: Very general pattern with N prefix - allowing for variable spaces
            /N\s*(\d{3,4})[\s\n]+([^\d]+?)[\s\n]+([\d.,]+)/g
        ];
        
        for (const pattern of patterns) {
            let match;
            let matchCount = 0;
            
            // Reset pattern index
            pattern.lastIndex = 0;
            
            while ((match = pattern.exec(lawSection)) !== null) {
                let nummer, bezeichnung, betrag, vorzeichen;
                
                if (pattern === patterns[0]) {
                    // Pattern 1 with codes
                    nummer = match[1];
                    bezeichnung = match[2].trim();
                    betrag = match[6].replace(/\./g, '').replace(',', '.');
                    vorzeichen = match[7] || '';
                } else if (pattern === patterns[1]) {
                    // Pattern 2 without codes
                    nummer = match[1];
                    bezeichnung = match[2].trim();
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = match[4] || '';
                } else {
                    // Pattern 3: Very general
                    nummer = match[1];
                    bezeichnung = match[2].trim();
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = '';
                }
                
                // Format the amount
                const betragValue = vorzeichen === '-' ? -parseFloat(betrag) : parseFloat(betrag);
                
                if (!isNaN(betragValue)) {
                    matchCount++;
                    result.push({
                        "Personalnummer": personalnummer,
                        "Lohnart": nummer,
                        "Lohnart Bezeichnung": "backdated (Lohnartenwerte)",
                        "Lohnart Betrag": betragValue
                    });
                }
            }
            
            // If we found matches with this pattern, no need to try others
            if (matchCount > 0) break;
        }
        
        return result;
    }
    
    // Extract regular NBA entries (without N prefix)
    function extractRegularNbaEntries(section, personalnummer) {
        const result = [];
        
        // Clean up whitespace
        section = section.replace(/\s+/g, ' ').trim();
        
        // Various patterns to try
        const patterns = [
            // Pattern 1: Standard format
            /\s(9\d{3})\s+(.{1,25}?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})(-?)/g,
            
            // Pattern 2: Without formatting
            /\s(9\d{3})\s+([^\d]+?)\s+([\d.,]+)(-?)/g,
            
            // Pattern 3: Very general pattern
            /\b(9\d{3})\b\s+([^\d]+?)\s+\b([\d.,]+)\b/g
        ];
        
        for (const pattern of patterns) {
            let match;
            let matchCount = 0;
            
            // Reset pattern index
            pattern.lastIndex = 0;
            
            while ((match = pattern.exec(section)) !== null) {
                const nummer = match[1];
                const bezeichnung = match[2].trim();
                let betrag, vorzeichen;
                
                if (pattern === patterns[0]) {
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = match[4] || '';
                } else if (pattern === patterns[1]) {
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = match[4] || '';
                } else {
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = '';
                }
                
                // Skip N-prefixed entries
                if (nummer.startsWith('N') || nummer.startsWith('n')) continue;
                
                // Filter out certain keywords
                const skipKeywords = ["Lohnsteuer", "Kirchensteuer", "Solidarit", "Steuerfreie", "P.", "aus NB"];
                if (skipKeywords.some(keyword => bezeichnung.includes(keyword))) continue;
                
                // Format the amount
                const betragValue = vorzeichen === '-' ? -parseFloat(betrag) : parseFloat(betrag);
                
                if (!isNaN(betragValue)) {
                    matchCount++;
                    result.push({
                        "Personalnummer": personalnummer,
                        "Netto-Bezüge/ -Abzüge Nummer": nummer,
                        "Netto-Bezüge/ -Abzüge Bezeichnung": bezeichnung,
                        "Netto-Bezüge/ -Abzüge Betrag": betragValue
                    });
                }
            }
            
            // If found matches with this pattern, no need to try others
            if (matchCount > 0) break;
        }
        
        return result;
    }
    
    // Extract N-prefixed NBA entries
    function extractNPrefixedNbaEntries(section, personalnummer) {
        const result = [];
        
        // Clean up whitespace
        section = section.replace(/\s+/g, ' ').trim();
        
        // Various patterns to try - with improved space handling between N and digits
        const patterns = [
            // Pattern 1: With N prefix - allowing for variable spaces
            /N\s*(9\d{3})\s+(.{1,25}?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})(-?)/g,
            
            // Pattern 2: With N prefix, without formatting - allowing for variable spaces
            /N\s*(9\d{3})\s+([^\d]+?)\s+([\d.,]+)(-?)/g,
            
            // Pattern 3: Very general pattern with N prefix - allowing for variable spaces
            /N\s*(9\d{3})[\s\n]+([^\d]+?)[\s\n]+([\d.,]+)/g
        ];
        
        for (const pattern of patterns) {
            let match;
            let matchCount = 0;
            
            // Reset pattern index
            pattern.lastIndex = 0;
            
            while ((match = pattern.exec(section)) !== null) {
                const nummer = match[1];
                const bezeichnung = match[2].trim();
                let betrag, vorzeichen;
                
                if (pattern === patterns[0]) {
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = match[4] || '';
                } else if (pattern === patterns[1]) {
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = match[4] || '';
                } else {
                    betrag = match[3].replace(/\./g, '').replace(',', '.');
                    vorzeichen = '';
                }
                
                // Filter out certain keywords
                const skipKeywords = ["Lohnsteuer", "Kirchensteuer", "Solidarit", "Steuerfreie", "P.", "aus NB"];
                if (skipKeywords.some(keyword => bezeichnung.includes(keyword))) continue;
                
                // Format the amount
                const betragValue = vorzeichen === '-' ? -parseFloat(betrag) : parseFloat(betrag);
                
                if (!isNaN(betragValue)) {
                    matchCount++;
                    result.push({
                        "Personalnummer": personalnummer,
                        "Netto-Bezüge/ -Abzüge Nummer": nummer,
                        "Netto-Bezüge/ -Abzüge Bezeichnung": "backdated (Lohnartenwerte)",
                        "Netto-Bezüge/ -Abzüge Betrag": betragValue
                    });
                }
            }
            
            // If found matches with this pattern, no need to try others
            if (matchCount > 0) break;
        }
        
        return result;
    }
    
    // Extract "already paid" entries
    function extractAlreadyPaidEntries(section, personalnummer) {
        const result = [];
        
        // Clean up whitespace
        section = section.replace(/\s+/g, ' ').trim();
        
        // Multiple patterns for already paid
        const patterns = [
            /bereits ausbezahlt(?:\s+([\d.,]+)(-?))?/g,
            /already paid(?:\s+([\d.,]+)(-?))?/g,
            /ausbezahlt(?:\s+([\d.,]+)(-?))?/g,
            /bezahlt(?:\s+([\d.,]+)(-?))?/g
        ];
        
        for (const pattern of patterns) {
            let match;
            let matchCount = 0;
            
            // Reset pattern index
            pattern.lastIndex = 0;
            
            while ((match = pattern.exec(section)) !== null) {
                if (match[1] && match[1].trim()) {
                    const betragStr = match[1];
                    const vorzeichen = match[2] || '';
                    
                    // Format the amount
                    const betragValue = vorzeichen === '-' ? 
                        -parseFloat(betragStr.replace(/\./g, '').replace(',', '.')) : 
                        parseFloat(betragStr.replace(/\./g, '').replace(',', '.'));
                    
                    if (!isNaN(betragValue)) {
                        matchCount++;
                        result.push({
                            "Personalnummer": personalnummer,
                            "Netto-Bezüge/ -Abzüge Nummer": "BEREITS_AUSBEZAHLT",
                            "Netto-Bezüge/ -Abzüge Bezeichnung": "already paid",
                            "Netto-Bezüge/ -Abzüge Betrag": betragValue
                        });
                    }
                }
            }
            
            // If found matches with this pattern, no need to try others
            if (matchCount > 0) break;
        }
        
        return result;
    }
    
    // Extract "aus NB" entries
    function extractAusNbEntries(section, personalnummer) {
        const result = [];
        
        // Clean up whitespace
        section = section.replace(/\s+/g, ' ').trim();
        
        // Multiple patterns for "aus NB" entries
        const patterns = [
            /aus NB\s+(.*?)\s+([\d.,]+)(-?)/g,
            /aus\s+NB\s+(.*?)[\s\n]+([\d.,]+)(-?)/g,
            /NB\s+(.*?)[\s\n]+([\d.,]+)(-?)/g
        ];
        
        for (const pattern of patterns) {
            let match;
            let matchCount = 0;
            
            // Reset pattern index
            pattern.lastIndex = 0;
            
            while ((match = pattern.exec(section)) !== null) {
                const bezeichnung = match[1].trim();
                const betragStr = match[2];
                const vorzeichen = match[3] || '';
                
                // Format the amount
                const betragValue = vorzeichen === '-' ? 
                    -parseFloat(betragStr.replace(/\./g, '').replace(',', '.')) : 
                    parseFloat(betragStr.replace(/\./g, '').replace(',', '.'));
                
                if (!isNaN(betragValue)) {
                    matchCount++;
                    result.push({
                        "Personalnummer": personalnummer,
                        "Netto-Bezüge/ -Abzüge Nummer": "aus_NB_" + bezeichnung,
                        "Netto-Bezüge/ -Abzüge Bezeichnung": "aus NB",
                        "Netto-Bezüge/ -Abzüge Betrag": betragValue
                    });
                }
            }
            
            // If found matches with this pattern, no need to try others
            if (matchCount > 0) break;
        }
        
        return result;
    }
    
    // Function to read Excel file
    async function readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = function(error) {
                reject(error);
            };
            
            reader.readAsArrayBuffer(file);
        });
    }
    
    // Function to merge data
    function mergeData(pdfData, lohnjournalData, personalkostenData, currentMonth) {
        // Merge the two Excel files first
        const merged_excel = mergeDataFrames(lohnjournalData, personalkostenData, 'Pers.Nr.');
        
        // Merge with calculated tax/social data
        const merged_df_all = mergeDataFrames(merged_excel, pdfData.tax_social_calculated, 'Pers.Nr.');
        
        // Log data for debugging
        console.log("Data to be merged:");
        console.log("- NBA data:", pdfData.nba_data.length, "entries");
        console.log("- LAW data:", pdfData.law_data.length, "entries");
        console.log("- Tax/Social data:", pdfData.tax_social_calculated.length, "entries");
        
        return {
            merged_df_all,
            combined_nba_df: pdfData.nba_data,
            combined_law_df: pdfData.law_data
        };
    }
    
    // Helper function to merge data frames
    function mergeDataFrames(df1, df2, on) {
        if (!df1 || !df1.length) return df2 || [];
        if (!df2 || !df2.length) return df1;
        
        const result = [];
        const df2ByKey = {};
        
        // Index df2 by the merge key
        df2.forEach(row => {
            if (row[on] !== undefined) {
                df2ByKey[row[on]] = row;
            }
        });
        
        // Merge df1 with df2
        df1.forEach(row1 => {
            if (row1[on] !== undefined && df2ByKey[row1[on]]) {
                const row2 = df2ByKey[row1[on]];
                const mergedRow = { ...row1 };
                
                // Add columns from df2 that aren't in df1
                Object.keys(row2).forEach(key => {
                    if (key !== on && mergedRow[key] === undefined) {
                        mergedRow[key] = row2[key];
                    }
                });
                
                result.push(mergedRow);
                delete df2ByKey[row1[on]]; // Remove processed entry
            } else {
                result.push(row1);
            }
        });
        
        // Add remaining rows from df2
        Object.values(df2ByKey).forEach(row2 => {
            result.push(row2);
        });
        
        return result;
    }
    
    // Function to create Excel file
    function createExcelFile(data, filename) {
        const workbook = XLSX.utils.book_new();
        
        // Get the special codes from the input fields to filter rows
        const bAVspecialcode = document.getElementById('bAVspecialcode').value;
        const NBAspecialcode = document.getElementById('NBAspecialcode').value;
        
        // Convert comma-separated special codes to arrays
        const bAVCodes = bAVspecialcode.split(',').map(code => code.trim()).filter(code => code !== '');
        const NBACodes = NBAspecialcode.split(',').map(code => code.trim()).filter(code => code !== '');
        
        console.log("Filtering out bAV codes:", bAVCodes);
        console.log("Filtering out NBA codes:", NBACodes);
        
        // Filter out rows with special codes from LAW data
        const filteredLawData = data.combined_law_df.filter(row => {
            // Keep the row only if its Lohnart is NOT in the bAV special codes list
            return !bAVCodes.includes(row.Lohnart);
        });
        
        // Filter out rows with special codes from NBA data
        const filteredNbaData = data.combined_nba_df.filter(row => {
            // Keep the row only if its Nummer is NOT in the NBA special codes list
            return !NBACodes.includes(row["Netto-Bezüge/ -Abzüge Nummer"]);
        });
        
        // Add sheets - merged data remains unchanged
        const ws1 = XLSX.utils.json_to_sheet(data.merged_df_all);
        XLSX.utils.book_append_sheet(workbook, ws1, 'Zusammengeführte Daten');
        
        // Add LAW data sheet (even if empty, with headers)
        if (filteredLawData.length > 0) {
            const ws2 = XLSX.utils.json_to_sheet(filteredLawData);
            XLSX.utils.book_append_sheet(workbook, ws2, 'Lohnartenwerte');
        } else {
            // Create empty sheet with headers
            const ws2 = XLSX.utils.aoa_to_sheet([["Personalnummer", "Lohnart", "Lohnart Bezeichnung", "Lohnart Betrag"]]);
            XLSX.utils.book_append_sheet(workbook, ws2, 'Lohnartenwerte');
        }
        
        // Add NBA data sheet (even if empty, with headers)
        if (filteredNbaData.length > 0) {
            const ws3 = XLSX.utils.json_to_sheet(filteredNbaData);
            XLSX.utils.book_append_sheet(workbook, ws3, 'NBA Daten');
        } else {
            // Create empty sheet with headers
            const ws3 = XLSX.utils.aoa_to_sheet([["Personalnummer", "Netto-Bezüge/ -Abzüge Nummer", "Netto-Bezüge/ -Abzüge Bezeichnung", "Netto-Bezüge/ -Abzüge Betrag"]]);
            XLSX.utils.book_append_sheet(workbook, ws3, 'NBA Daten');
        }
        
        // Generate Excel file
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }
});