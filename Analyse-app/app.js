// Excel Statistics Application
class ExcelStatsApp {
    constructor() {
        this.rawData = null;
        this.processedData = null; 
        this.headers = [];
        this.numericColumns = [];
        this.selectedColumnsIndices = []; 
        this.statistics = {}; 
        this.charts = {}; 
        this.currentFile = null; 
        this.logoDataUrl = null; 
        this.studentQuartileCategories = { q1Students: [], q2Students: [], q3Students: [], q4Students: [] };
        this.quartileValues = { q1: null, median: null, q3: null }; 
        this.gradeParsingMode = 'decimal'; // Default: +/- gir desimaler
        this.initializeEventListeners();
        this.setGlobalChartStyles();
    }

    setGlobalChartStyles() {
        try {
            const rootStyle = getComputedStyle(document.documentElement);
            const textColor = rootStyle.getPropertyValue('--color-text').trim();
            const borderColor = rootStyle.getPropertyValue('--color-border').trim();
            const fontFamily = rootStyle.getPropertyValue('--font-family-base').trim() || 'sans-serif';

            if (textColor) Chart.defaults.color = textColor;
            if (borderColor) Chart.defaults.borderColor = borderColor;
            Chart.defaults.font.family = fontFamily;
            
        } catch (e) {
            console.warn("Kunne ikke sette globale Chart.js farger fra CSS-variabler.", e);
            Chart.defaults.color = '#363636'; 
            Chart.defaults.borderColor = '#e0e0e0'; 
            Chart.defaults.font.family = 'Helvetica, Arial, sans-serif';
        }
    }

    initializeEventListeners() {
        const fileInput = document.getElementById('file-input');
        const uploadArea = document.getElementById('upload-area');
        const logoInput = document.getElementById('logo-input'); 
        const plusMinusHandlingSelect = document.getElementById('plus-minus-handling');

        if (fileInput) {
            fileInput.addEventListener('change', (e) => {
                if (e.target.files.length > 0) {
                    this.currentFile = e.target.files[0];
                    this.handleFileSelect(this.currentFile); 
                }
            });
        } else { console.error("Element 'file-input' not found."); }

        if (uploadArea) {
            uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.classList.add('drag-over'); });
            uploadArea.addEventListener('dragleave', (e) => { e.preventDefault(); uploadArea.classList.remove('drag-over'); });
            uploadArea.addEventListener('drop', (e) => {
                e.preventDefault();
                uploadArea.classList.remove('drag-over');
                if (e.dataTransfer.files.length > 0) {
                    this.currentFile = e.dataTransfer.files[0];
                    this.handleFileSelect(this.currentFile); 
                }
            });
        } else { console.error("Element 'upload-area' not found."); }


        if (logoInput) {
            logoInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                const logoPreview = document.getElementById('logo-preview');
                if (file) {
                    if (file.size > 2 * 1024 * 1024) { 
                        this.showError("Logofilen er for stor (maks 2MB).");
                        logoInput.value = null; 
                        if(logoPreview) logoPreview.classList.add('hidden');
                        this.logoDataUrl = null;
                        return;
                    }
                    const reader = new FileReader();
                    reader.onload = (event) => {
                        this.logoDataUrl = event.target.result;
                        if (logoPreview) {
                            logoPreview.src = this.logoDataUrl;
                            logoPreview.classList.remove('hidden');
                        }
                        this.showInfo("Logo lastet opp. Den vil bli inkludert i PDF-rapportene.");
                    };
                    reader.onerror = () => {
                        this.showError("Kunne ikke lese logofilen.");
                        this.logoDataUrl = null;
                        if(logoPreview) logoPreview.classList.add('hidden');
                    };
                    reader.readAsDataURL(file);
                } else {
                    this.logoDataUrl = null; 
                    if (logoPreview) logoPreview.classList.add('hidden');
                }
            });
        } else { console.warn("Element 'logo-input' not found (optional).");}

        if (plusMinusHandlingSelect) { 
            plusMinusHandlingSelect.addEventListener('change', (e) => {
                this.gradeParsingMode = e.target.value;
                console.log("Karaktertolkningsmodus endret til:", this.gradeParsingMode);
                if (this.processedData && this.currentFile) { // Hvis data er lastet og analysert
                    this.showInfo("Karaktertolkningsmodus endret. Kjør analysen på nytt for at endringen skal tre i kraft.");
                    // Vurder å nullstille `processedData` og tvinge re-analyse, eller la brukeren gjøre det.
                    // For nå, bare informer.
                }
            });
        } else { console.warn("Element 'plus-minus-handling' select not found.");}


        const analyzeBtn = document.getElementById('analyze-btn');
        if(analyzeBtn) analyzeBtn.addEventListener('click', () => this.analyzeData());
        else console.error("Element 'analyze-btn' not found.");

        const generatePdfsBtn = document.getElementById('generate-pdfs-btn');
        if(generatePdfsBtn) generatePdfsBtn.addEventListener('click', () => this.generateAllStudentPDFs());
        else console.error("Element 'generate-pdfs-btn' not found.");
        
        const closePdfLinksBtn = document.getElementById('close-pdf-links-btn');
        if (closePdfLinksBtn) {
            closePdfLinksBtn.addEventListener('click', () => {
                const pdfContainer = document.getElementById('pdf-download-links-container');
                if (pdfContainer) pdfContainer.classList.add('hidden');
            });
        } else { console.warn("Element 'close-pdf-links-btn' not found (optional).");}

        document.body.addEventListener('click', async (event) => {
            const target = event.target.closest('.individual-pdf-download-btn');
            if (target && target.dataset.studentIndex !== undefined) {
                const studentIndex = parseInt(target.dataset.studentIndex, 10);
                if (this.processedData && this.processedData[studentIndex]) {
                    const student = this.processedData[studentIndex];
                    this.showInfo(`Forbereder PDF for ${student.name}...`);
                    await this.downloadSingleStudentPDF(student);
                } else {
                    this.showError("Kunne ikke finne studentdata for den valgte indeksen.");
                    console.error("Ugyldig studentindeks for PDF-nedlasting:", studentIndex, this.processedData);
                }
            }
        });
    }

    async handleFileSelect(file) { 
        if (file) {
            const resultsSection = document.getElementById('results-section');
            if(resultsSection) resultsSection.classList.add('hidden');
            const configSection = document.getElementById('config-section');
            if(configSection) configSection.classList.add('hidden');
            const pdfLinksContainer = document.getElementById('pdf-download-links-container');
            if(pdfLinksContainer) pdfLinksContainer.classList.add('hidden'); 
            
            this.rawData = null;
            this.processedData = null;
            this.destroyAllCharts(); 
            this.charts = {}; 
            this.headers = [];
            this.numericColumns = [];
            this.selectedColumnsIndices = [];
            this.statistics = {};
            this.studentQuartileCategories = { q1Students: [], q2Students: [], q3Students: [], q4Students: [] };
            this.quartileValues = { q1: null, median: null, q3: null };

            const dynamicAreas = ['column-checkboxes', 'stats-table-body', 'heatmap-chart-table', 
                                  'student-sparklines', 'pdf-links-list', 
                                  'quartile1-list', 'quartile2-list', 'quartile3-list', 'quartile4-list',
                                  'quartile-info-text'];
            dynamicAreas.forEach(id => {
                const el = document.getElementById(id);
                if (el) el.innerHTML = '';
                else console.warn(`Dynamisk område med ID '${id}' ikke funnet under reset.`);
            });
            
            await this.processFile(file);
        }
    }

    async processFile(file) {
        this.showProgress(true, 'Leser fil...');
        this.updateProgress(0, 'Starter filbehandling...');
        try {
            const plusMinusSelect = document.getElementById('plus-minus-handling');
            if (plusMinusSelect) {
                this.gradeParsingMode = plusMinusSelect.value;
            } else {
                this.gradeParsingMode = 'decimal'; 
            }
            console.log("Bruker karaktertolkningsmodus:", this.gradeParsingMode, "for behandling av filen:", file.name);

            await new Promise(resolve => setTimeout(resolve, 50)); 
            this.updateProgress(25, `Leser filstruktur fra "${file.name}"...`);
            const data = await this.readExcelFile(file);
            this.updateProgress(50, 'Behandler rådata...');
            this.rawData = data;
            if (!this.rawData || this.rawData.length === 0) { throw new Error("Excel-filen er tom eller inneholder ikke gjenkjennbare data."); }
            this.headers = this.rawData[0] ? this.rawData[0].map(header => String(header || '').trim()) : [];
            if (this.headers.length === 0) { throw new Error("Fant ingen kolonneoverskrifter. Sjekk at første rad inneholder navn på kolonnene."); }
            this.updateProgress(75, 'Identifiserer karakterkolonner...');
            this.identifyNumericColumns(); 
            if (this.numericColumns.length === 0) { this.showInfo("Ingen kolonner med gjenkjennelige karakterer funnet. Sjekk formatet i Excel-filen."); }
            this.updateProgress(100, 'Filbehandling fullført.');
            setTimeout(() => { 
                this.showProgress(false);
                this.showConfigSection();
                this.populateColumnCheckboxes();
                this.showSuccess(`Filen "${file.name}" (${(file.size / 1024).toFixed(1)}KB) er lastet og klar for konfigurasjon.`);
            }, 600);
        } catch (error) {
            console.error('Error processing file:', error);
            this.showError(`Feil ved filbehandling: ${error.message}`);
            this.showProgress(false);
            const fileInput = document.getElementById('file-input');
            if(fileInput) fileInput.value = null; 
        }
    }

    async readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const fileData = new Uint8Array(e.target.result);
                    const parseStrategies = [
                        { opts: { type: 'array', cellText: true, cellNF: false, cellHTML: false, raw: false, defval: null }, msg: "Default array strategy" },
                        { opts: { type: 'array', raw: true, cellText: false, defval: null }, msg: "Raw array strategy" },
                        { binary: true, opts: { type: 'binary', cellText: true, raw: false, defval: null }, msg: "Binary string strategy" }
                    ];

                    for (let strategyInfo of parseStrategies) {
                        try {
                            console.log(`Prøver lesestrategi: ${strategyInfo.msg}`);
                            let workbook;
                            if (strategyInfo.binary) {
                                const binaryString = Array.from(fileData).map(byte => String.fromCharCode(byte)).join('');
                                workbook = XLSX.read(binaryString, strategyInfo.opts);
                            } else {
                                workbook = XLSX.read(fileData, strategyInfo.opts);
                            }
                            const result = this.extractWorksheetData(workbook);
                            if (result && result.length > 0) {
                                console.log(`Suksess med strategi: ${strategyInfo.msg}`);
                                resolve(result);
                                return;
                            }
                        } catch (strategyError) {
                            console.warn(`Lesestrategi "${strategyInfo.msg}" feilet:`, strategyError.message);
                        }
                    }
                    reject(new Error('Kunne ikke lese Excel-filen. Filen kan være korrupt eller i et format som ikke støttes. Prøv å lagre som .xlsx.'));
                } catch (error) {
                    reject(new Error(`Ugyldig Excel-fil format eller intern feil: ${error.message}`));
                }
            };
            reader.onerror = (err) => reject(new Error(`En feil oppstod under lesing av filen: ${err.type || 'Ukjent lesefeil'}`));
            reader.readAsArrayBuffer(file);
        });
    }

    extractWorksheetData(workbook) {
        const sheetNames = workbook.SheetNames;
        if (!sheetNames || sheetNames.length === 0) {
            throw new Error("Ingen regneark (arkfaner) funnet i Excel-filen.");
        }
        
        for (let sheetName of sheetNames) {
            try {
                const worksheet = workbook.Sheets[sheetName];
                if (!worksheet || !worksheet['!ref']) continue; 

                let data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: null, blankrows: false });
                
                if (data && data.length > 0) {
                    const filteredData = data.filter(row => 
                        Array.isArray(row) && row.some(cell => cell !== null && String(cell).trim() !== "")
                    );
                    if (filteredData.length > 0) {
                        if (filteredData[0] && filteredData[0].some(header => header !== null && String(header).trim() !== "")) {
                             console.log(`Hentet ut ${filteredData.length} rader fra arket "${sheetName}".`);
                            return filteredData;
                        }
                    }
                }
            } catch (sheetError) {
                console.warn(`Feilet å hente data fra arket "${sheetName}":`, sheetError.message);
            }
        }
        throw new Error('Kunne ikke hente ut gyldige data fra noen regneark. Sørg for at arket inneholder data og at første rad har kolonneoverskrifter.');
    }
    
    identifyNumericColumns() {
        this.numericColumns = [];
        if (!this.rawData || this.rawData.length < 2 || !this.headers || this.headers.length === 0) {
            console.warn("Kan ikke identifisere numeriske kolonner: utilstrekkelig rådata.");
            return;
        }
        const dataRows = this.rawData.slice(1); 
        this.headers.forEach((headerName, colIndex) => {
            if (colIndex === 0 || !headerName || String(headerName).trim() === "") return; 
            let numericCellCount = 0;
            let nonEmptyCellCount = 0; 
            for (let row of dataRows) {
                if (row && row.length > colIndex) { 
                    const value = row[colIndex];
                    if (value !== null && String(value).trim() !== "") {
                        nonEmptyCellCount++;
                        if (this.processNorwegianGrade(value) !== null) { 
                            numericCellCount++;
                        }
                    }
                }
            }
            if (nonEmptyCellCount >= 3 && (numericCellCount / nonEmptyCellCount) >= 0.5) {
                this.numericColumns.push({ name: String(headerName), index: colIndex });
            }
        });
        console.log(`Identifiserte karakterkolonner (med tolkningsmodus '${this.gradeParsingMode}'):`, this.numericColumns.map(c => c.name));
    }

    processNorwegianGrade(value) {
        if (value === null || value === undefined) return null;
        let str = String(value).trim();
        if (str === "") return null;

        const nonGradeTexts = [
            "iv", "ikke vurdert", "im", "fritatt", "godkjent", "bestått", "deltatt", 
            "g", "mg", "ng", "ig", "ikke møtt", "syk", "permisjon", "gyldig fravær",
            "levert", "ikke levert", "vurdering kommer", "muntlig", "skriftlig"
        ];
        const lowerStr = str.toLowerCase();
        if (nonGradeTexts.some(term => lowerStr.includes(term))) return null;

        const pureNumericMatch = str.match(/^([1-6])([,\.]\d{1,2})?$/);
        if (pureNumericMatch) {
            const numericValue = parseFloat(pureNumericMatch[0].replace(',', '.'));
            if (numericValue >= 1 && numericValue <= 6) return parseFloat(numericValue.toFixed(1));
        }
        
        const plusMatch = str.match(/^([1-5])\+$/);
        if (plusMatch) {
            const baseGrade = parseInt(plusMatch[1]);
            return this.gradeParsingMode === 'decimal' ? baseGrade + 0.3 : baseGrade;
        }

        const minusMatch = str.match(/^([2-6])\-$/);
        if (minusMatch) {
            const baseGrade = parseInt(minusMatch[1]);
            return this.gradeParsingMode === 'decimal' ? baseGrade - 0.3 : baseGrade;
        }
        
        const slashGradeMatch = str.match(/^([1-6])([,\.]\d)?\s*\/\s*([1-6])([,\.]\d)?$/); 
        if (slashGradeMatch) {
             const g1Str = slashGradeMatch[1] + (slashGradeMatch[2] || '').replace(',', '.');
             const g2Str = slashGradeMatch[3] + (slashGradeMatch[4] || '').replace(',', '.');
             const g1 = parseFloat(g1Str);
             const g2 = parseFloat(g2Str);
             if(!isNaN(g1) && !isNaN(g2)) return parseFloat(((g1 + g2) / 2).toFixed(1)); 
        }

        const parsedFloatSimple = parseFloat(str.replace(',', '.'));
        if (!isNaN(parsedFloatSimple) && parsedFloatSimple >= 1 && parsedFloatSimple <= 6) {
             return parseFloat(parsedFloatSimple.toFixed(1));
        }

        if (isNaN(parseFloat(str)) && str.length > 2 && !str.match(/[1-6]/)) { 
            // console.log(`Ignorerer verdi som tekstkommentar: "${str}" (Modus: ${this.gradeParsingMode})`);
            return null; 
        }
        
        return null; 
    }


    populateColumnCheckboxes() {
        const container = document.getElementById('column-checkboxes');
        if(!container) { console.error("Element 'column-checkboxes' not found for populating."); return; }
        container.innerHTML = ''; 
        if (this.numericColumns.length === 0) {
            container.innerHTML = '<p style="color: var(--color-text-secondary);">Ingen kolonner med gjenkjennelige karakterer ble funnet. Sjekk Excel-filen og valgt tolkningsmodus for +/-.</p>';
            return;
        }
        this.numericColumns.forEach(column => {
            const div = document.createElement('div');
            div.className = 'checkbox-item';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `col-check-${column.index}`; 
            checkbox.value = column.index;
            checkbox.checked = true; 
            const label = document.createElement('label');
            label.htmlFor = `col-check-${column.index}`;
            label.textContent = column.name;
            div.appendChild(checkbox);
            div.appendChild(label);
            container.appendChild(div);
        });
    }

    analyzeData() {
        const plusMinusSelect = document.getElementById('plus-minus-handling');
        if (plusMinusSelect) {
            this.gradeParsingMode = plusMinusSelect.value;
        } else {
            this.gradeParsingMode = 'decimal'; 
        }
        console.log("Analyserer data med karaktertolkningsmodus:", this.gradeParsingMode);

        this.showProgress(true, 'Analyserer data...');
        this.updateProgress(0, 'Starter analyse...');
        try {
            const checkboxes = document.querySelectorAll('#column-checkboxes input[type="checkbox"]:checked');
            this.selectedColumnsIndices = Array.from(checkboxes).map(cb => parseInt(cb.value)); 
            if (this.selectedColumnsIndices.length === 0) {
                this.showError('Ingen vurderingskolonner er valgt.');
                this.showProgress(false);
                return;
            }
            this.updateProgress(10, 'Henter konfigurasjonsdata...');
            const schoolName = document.getElementById('school-name').value.trim() || 'Ukjent skole';
            const subject = document.getElementById('subject').value.trim() || 'Ukjent fag';
            const schoolYear = document.getElementById('school-year').value.trim() || 'Ukjent skoleår/periode';

            this.updateProgress(25, 'Behandler elevdata...');
            this.processStudentDataAndCalculateAverages(); 
            if(!this.processedData || this.processedData.length === 0){
                this.showInfo("Ingen gyldige elevdata funnet etter behandling.");
                this.showProgress(false);
                return;
            }
            
            this.updateProgress(40, 'Kategoriserer elever...');
            this.calculateQuartilesAndCategorizeStudents();

            this.updateProgress(50, 'Kalkulerer statistikk...');
            this.calculateOverallStatistics(); 

            this.updateProgress(75, 'Forbereder resultater...');
            this.displayResults(schoolName, subject, schoolYear); 
            this.createAllVisualizations(); 
            
            this.updateProgress(100, 'Analyse fullført.');
            const resultsSection = document.getElementById('results-section');
            if (resultsSection) {
                resultsSection.classList.remove('hidden');
                resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
            this.showSuccess('Dataanalyse fullført!');
            this.showProgress(false);

        } catch (error) {
            console.error('Error analyzing data:', error);
            this.showError(`En uventet feil oppstod under analysen: ${error.message}`);
            this.showProgress(false);
        }
    }

    processStudentDataAndCalculateAverages() { 
        if (!this.rawData || this.rawData.length < 2) { this.processedData = []; return; }
        this.headers = this.rawData[0] ? this.rawData[0].map(header => String(header || '').trim()) : [];
        if (this.headers.length === 0) { this.processedData = []; return; }

        console.log(`processStudentDataAndCalculateAverages (re-check) kjører med gradeParsingMode: ${this.gradeParsingMode}`);

        this.processedData = this.rawData.slice(1).map((row, studentIndexInRaw) => { 
            const studentName = (row && row[0] !== null && String(row[0]).trim() !== "") ? String(row[0]).trim() : null; 
            if (!studentName) return null; 

            const processedRow = { name: studentName, originalIndexInRawData: studentIndexInRaw + 1 }; 
            let sumOfGrades = 0;
            let countOfGrades = 0;
            
            this.selectedColumnsIndices.forEach(colIndex => { 
                if (this.headers[colIndex] !== undefined) { 
                    const columnName = this.headers[colIndex];
                    const rawValue = (row && row.length > colIndex) ? row[colIndex] : null;
                    const grade = this.processNorwegianGrade(rawValue); 
                    processedRow[columnName] = grade; 
                    if (grade !== null && !isNaN(grade)) { 
                        sumOfGrades += grade;
                        countOfGrades++;
                    }
                }
            });
            processedRow.averageGrade = countOfGrades > 0 ? parseFloat((sumOfGrades / countOfGrades).toFixed(1)) : null;
            
            // Detaljert logging per student for snittberegning (kan aktiveres ved behov)
            // if(studentName === "EN SPESIFIKK ELEV") { // Bytt ut med et reelt navn for feilsøking
            //    const gradesForAvgLog = [];
            //    this.selectedColumnsIndices.forEach(colIndex => {
            //        const rawValue = (row && row.length > colIndex) ? row[colIndex] : null;
            //        const gradeVal = this.processNorwegianGrade(rawValue);
            //        gradesForAvgLog.push({column: this.headers[colIndex], raw: rawValue, processed: gradeVal, included: (gradeVal !== null && !isNaN(gradeVal))});
            //    });
            //    console.log(`--- Snittberegning for ${studentName} (Modus: ${this.gradeParsingMode}) ---`);
            //    console.log("Valgte kolonner:", this.selectedColumnsIndices.map(i => `${this.headers[i]} (idx ${i})`).join(', '));
            //    console.log("Karakterdetaljer:", gradesForAvgLog);
            //    console.log(`Sum: ${sumOfGrades.toFixed(2)}, Antall: ${countOfGrades}, Beregnet Snitt (student.averageGrade): ${processedRow.averageGrade}`);
            //    console.log("-----------------------------------------");
            // }

            return processedRow;
        }).filter(student => student !== null); 
        
        this.processedData.forEach((student, indexInProcessed) => {
            student.processedDataIndex = indexInProcessed; 
        });
        console.log(`Behandlet ${this.processedData.length} elever. student.averageGrade er beregnet.`);
    }
    
    getPercentileValue(sortedArray, percentile) {
        if (!sortedArray || sortedArray.length === 0) return null;
        const k = (sortedArray.length - 1) * (percentile / 100);
        const f = Math.floor(k);
        const c = Math.ceil(k);
        if (f === c) return sortedArray[f];
        const valF = sortedArray[Math.max(0, Math.min(f, sortedArray.length - 1))];
        const valC = sortedArray[Math.max(0, Math.min(c, sortedArray.length - 1))];
        return parseFloat(((valF * (c - k)) + (valC * (k - f))).toFixed(1));
    }


    calculateQuartilesAndCategorizeStudents() {
        this.studentQuartileCategories = { q1Students: [], q2Students: [], q3Students: [], q4Students: [] }; 
        this.quartileValues = { q1: null, median: null, q3: null };
        const quartileInfoEl = document.getElementById('quartile-info-text');

        if (!this.processedData || this.processedData.length === 0) {
            if(quartileInfoEl) quartileInfoEl.textContent = "Ingen data å kategorisere.";
            return;
        }
        const validAverageGrades = this.processedData.map(s => s.averageGrade).filter(avg => avg !== null && !isNaN(avg)).sort((a,b) => a-b);

        if (validAverageGrades.length < 4) { 
            this.showInfo("For få elever for kvartilinndeling. Grupperer som 'lav', 'middels', 'høy'.");
            const lowMax = 2.69, mediumMax = 4.49; 
            this.processedData.forEach(student => {
                if (student.averageGrade !== null && !isNaN(student.averageGrade)) {
                    if (student.averageGrade <= lowMax) this.studentQuartileCategories.q1Students.push(student);
                    else if (student.averageGrade <= mediumMax) this.studentQuartileCategories.q2Students.push(student);
                    else this.studentQuartileCategories.q4Students.push(student);
                }
            });
            if(quartileInfoEl) quartileInfoEl.textContent = "Enkel gruppering grunnet få elever.";
            return; 
        }
        
        this.quartileValues.q1 = this.getPercentileValue(validAverageGrades, 25);
        this.quartileValues.median = this.getPercentileValue(validAverageGrades, 50);
        this.quartileValues.q3 = this.getPercentileValue(validAverageGrades, 75);
        
        this.processedData.forEach(student => {
            if (student.averageGrade !== null && !isNaN(student.averageGrade)) {
                if (student.averageGrade < this.quartileValues.q1) student.category = "Laveste kvartil", this.studentQuartileCategories.q1Students.push(student);
                else if (student.averageGrade < this.quartileValues.median) student.category = "Nedre midtre kvartil", this.studentQuartileCategories.q2Students.push(student);
                else if (student.averageGrade < this.quartileValues.q3) student.category = "Øvre midtre kvartil", this.studentQuartileCategories.q3Students.push(student);
                else student.category = "Høyeste kvartil", this.studentQuartileCategories.q4Students.push(student);
            } else student.category = "Ukategorisert";
        });
        if(quartileInfoEl) quartileInfoEl.textContent = `Kvartilgrenser: Q1=${this.quartileValues.q1}, M=${this.quartileValues.median}, Q3=${this.quartileValues.q3}.`;
    }

    calculateOverallStatistics() { 
        this.statistics = {}; 
        this.selectedColumnsIndices.forEach(colIndex => {
            const columnName = this.headers[colIndex];
            const values = this.processedData.map(row => row[columnName]).filter(val => val !== null && !isNaN(val)); 
            if (values.length > 0) {
                const sum = values.reduce((a, b) => a + b, 0);
                const mean = sum / values.length;
                const sorted = [...values].sort((a, b) => a - b);
                const median = sorted.length % 2 === 0 ? (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2 : sorted[Math.floor(sorted.length / 2)];
                const variance = values.length > 1 ? values.reduce((sq, n) => sq + Math.pow(n - mean, 2), 0) / (values.length - 1) : 0;
                const std = values.length > 1 ? Math.sqrt(variance) : 0; 
                this.statistics[columnName] = { mean, median, std, min: Math.min(...values), max: Math.max(...values), count: values.length, values };
            } else this.statistics[columnName] = { mean:0, median:0, std:0, min:NaN, max:NaN, count:0, values:[] };
        });

        let studentsWithImprovement = 0, studentsWithMultipleValidGrades = 0; 
        if (this.selectedColumnsIndices.length >= 2) {
            const firstColName = this.headers[this.selectedColumnsIndices[0]];
            const lastColName = this.headers[this.selectedColumnsIndices[this.selectedColumnsIndices.length - 1]];
            this.processedData.forEach(student => {
                const firstGrade = student[firstColName], lastGrade = student[lastColName];
                if (firstGrade !== null && !isNaN(firstGrade) && lastGrade !== null && !isNaN(lastGrade)) {
                    studentsWithMultipleValidGrades++;
                    if (lastGrade > firstGrade) studentsWithImprovement++;
                }
            });
        }
        this.gradeImprovementPercentage = studentsWithMultipleValidGrades > 0 ? Math.round((studentsWithImprovement / studentsWithMultipleValidGrades) * 100) : 0;
        const allGradesFlat = Object.values(this.statistics).flatMap(stat => stat.values || []).filter(val => val !== null && !isNaN(val));
        this.performanceLevelsCounts = { 
            high: allGradesFlat.filter(g => g >= 5).length,
            medium: allGradesFlat.filter(g => g >= 3 && g < 5).length,
            low: allGradesFlat.filter(g => g < 2).length, 
            pass: allGradesFlat.filter(g => g >= 2).length 
        };
        this.passRate = allGradesFlat.length > 0 ? Math.round((this.performanceLevelsCounts.pass / allGradesFlat.length) * 100) : 0;
        this.averageGradeOverall = allGradesFlat.length > 0 ? (allGradesFlat.reduce((a,b) => a+b, 0) / allGradesFlat.length).toFixed(1) : '0.0';
    }

    safeSetTextContent(elementId, text) {
        const element = document.getElementById(elementId);
        if (element) element.textContent = String(text); 
        else console.warn(`Element med ID '${elementId}' ikke funnet.`);
    }

    displayResults(schoolName, subject, schoolYear) {
        this.safeSetTextContent('display-school', schoolName);
        this.safeSetTextContent('display-subject', subject);
        this.safeSetTextContent('display-year', schoolYear);
        this.safeSetTextContent('total-students', this.processedData ? this.processedData.length : 0);
        this.safeSetTextContent('average-grade-overall', this.averageGradeOverall); 
        this.safeSetTextContent('assessments-count', this.selectedColumnsIndices.length);
        this.safeSetTextContent('pass-rate', `${this.passRate}%`);
        this.safeSetTextContent('grade-improvement-percentage', `${this.gradeImprovementPercentage}%`); 
        this.safeSetTextContent('high-performers-count', this.performanceLevelsCounts.high); 
        this.safeSetTextContent('medium-performers-count', this.performanceLevelsCounts.medium);
        this.safeSetTextContent('low-performers-count', this.performanceLevelsCounts.low);
        this.displayStatisticsTable();
        this.displayStudentQuartileCategories(); 
    }

    displayStatisticsTable() {
        const tbody = document.getElementById('stats-table-body');
        if (!tbody) { console.error("Element 'stats-table-body' not found."); return; }
        tbody.innerHTML = '';
        const statNames = ['Gj.snitt', 'Median', 'Std.avvik', 'Laveste', 'Høyeste', 'Antall'];
        const statKeys = ['mean', 'median', 'std', 'min', 'max', 'count'];
        const headerRow = tbody.insertRow();
        let th = headerRow.insertCell(); th.textContent = 'Statistikk'; th.className = 'student-name-header'; 
        this.selectedColumnsIndices.forEach(idx => { th = headerRow.insertCell(); th.textContent = this.headers[idx]; th.className = 'assessment-header'; });
        statKeys.forEach((key, i) => {
            const row = tbody.insertRow(); let cell = row.insertCell(); cell.textContent = statNames[i]; cell.className = 'stat-row student-name-cell'; 
            this.selectedColumnsIndices.forEach(idx => {
                const stat = this.statistics[this.headers[idx]]; cell = row.insertCell(); cell.className = 'stat-value';
                if (stat && stat.count > 0) cell.textContent = (key==='count' || key==='min' || key==='max') ? stat[key].toFixed(key==='count'?0:1) : stat[key].toFixed(2);
                else if (key==='count' && stat) cell.textContent = '0'; else cell.textContent = '–';
            });
        });
    }

    displayStudentQuartileCategories() {
        const lists = {q1:document.getElementById('quartile1-list'), q2:document.getElementById('quartile2-list'), q3:document.getElementById('quartile3-list'), q4:document.getElementById('quartile4-list')};
        const quartileInfoEl = document.getElementById('quartile-info-text');
        Object.values(lists).forEach(list => { if(list) list.innerHTML = '';});
        const categoryMap = {q1:this.studentQuartileCategories.q1Students, q2:this.studentQuartileCategories.q2Students, q3:this.studentQuartileCategories.q3Students, q4:this.studentQuartileCategories.q4Students};
        const useFallback = this.processedData && this.processedData.length < 4 && this.quartileValues.q1 === null;

        for (const key in categoryMap) {
            const studentList = categoryMap[key]; const ulElement = lists[key.slice(0,2)];
            if (ulElement && studentList.length > 0) {
                studentList.sort((a,b)=>(a.averageGrade||0)-(b.averageGrade||0)).forEach(s=>{
                    const li=document.createElement('li');
                    li.innerHTML = `<span class="student-category-name">${s.name}</span><span class="student-category-avg">Snitt: ${s.averageGrade!==null?s.averageGrade.toFixed(1):'N/A'}</span><button class="individual-pdf-download-btn" data-student-index="${s.processedDataIndex}" title="PDF">📄</button>`;
                    ulElement.appendChild(li);
                });
            } else if (ulElement && !useFallback && this.quartileValues.q1 !== null) ulElement.innerHTML = '<li>Ingen elever.</li>';
        }
        const q1h5=document.getElementById('q1-threshold-display')?.parentElement?.querySelector('h5'), q2h5=document.getElementById('q2-threshold-display')?.parentElement?.querySelector('h5'), q3col=document.getElementById('q3-threshold-display')?.parentElement?.parentElement, q3h5=document.getElementById('q3-threshold-display')?.parentElement?.querySelector('h5'), q4h5=document.getElementById('q3-upper-threshold-display')?.parentElement?.querySelector('h5');
        if(useFallback){if(q1h5)q1h5.innerHTML=`Lav (<span class="quartile-threshold">&lt;2.7</span>)`; if(q2h5)q2h5.innerHTML=`Middels (<span class="quartile-threshold">2.7–4.4</span>)`; if(q3col)q3col.style.display='none'; if(q4h5)q4h5.innerHTML=`Høy (<span class="quartile-threshold">≥4.5</span>)`; if(quartileInfoEl)quartileInfoEl.textContent="Enkel gruppering.";}
        else if(this.quartileValues.q1!==null){if(q3col)q3col.style.display='block'; if(q1h5)q1h5.innerHTML=`Laveste kv. (<span class="quartile-threshold">&lt;${this.quartileValues.q1}</span>)`; if(q2h5)q2h5.innerHTML=`N. midtre (<span class="quartile-threshold">${this.quartileValues.q1}–&lt;${this.quartileValues.median}</span>)`; if(q3h5)q3h5.innerHTML=`Ø. midtre (<span class="quartile-threshold">${this.quartileValues.median}–&lt;${this.quartileValues.q3}</span>)`; if(q4h5)q4h5.innerHTML=`Høyeste kv. (<span class="quartile-threshold">≥${this.quartileValues.q3}</span>)`; if(quartileInfoEl)quartileInfoEl.textContent=`Kvartiler: Q1=${this.quartileValues.q1}, M=${this.quartileValues.median}, Q3=${this.quartileValues.q3}.`;}
        else if(quartileInfoEl)quartileInfoEl.textContent="Analyser data for kvartilinndeling.";
    }

    createAllVisualizations() {
        this.destroyAllCharts(); 
        this.createDistributionChart();
        this.createHeatmapTable(); 
        this.createSparklinesWithIndividualDownload(); 
        this.createClassProgressChart();
        this.createLevelGroupsChart();
    }
    
    destroyAllCharts() {
        Object.keys(this.charts).forEach(key => { if (this.charts[key]?.destroy) this.charts[key].destroy(); });
        this.charts = {}; 
    }

    getCssVar(varName) { try { return getComputedStyle(document.documentElement).getPropertyValue(varName).trim() || null; } catch (e) { return null; } }
    getCssVarWithAlpha(varName, alpha) {
        const color = this.getCssVar(varName); if (!color) return `rgba(33,128,141,${alpha})`;
        if(color.startsWith('rgba')) return color.replace(/,\s*\d?\.?\d*\s*\)$/,`, ${alpha})`);
        if(color.startsWith('rgb')) return color.replace(')',`, ${alpha})`).replace('rgb','rgba');
        if(color.startsWith('#')){let r=0,g=0,b=0; if(color.length===4){r=parseInt(color[1]+color[1],16);g=parseInt(color[2]+color[2],16);b=parseInt(color[3]+color[3],16);}else if(color.length===7){r=parseInt(color.substring(1,3),16);g=parseInt(color.substring(3,5),16);b=parseInt(color.substring(5,7),16);} return `rgba(${r},${g},${b},${alpha})`;}
        return `rgba(33,128,141,${alpha})`; 
    }

    createClassProgressChart() {
        const ctx = document.getElementById('class-progress-chart')?.getContext('2d'); if (!ctx) return;
        const averages = this.selectedColumnsIndices.map(idx => this.statistics[this.headers[idx]]?.mean ?? null);
        const labels = this.selectedColumnsIndices.map(idx => this.headers[idx]);
        const color = this.getCssVar('--color-chart-line-blue') || '#36A2EB', bgColor = this.getCssVarWithAlpha('--color-chart-line-blue', 0.1);
        if (this.charts.classProgress) this.charts.classProgress.destroy(); 
        this.charts.classProgress = new Chart(ctx, {type:'line', data:{labels, datasets:[{label:'Klassesnitt',data:averages,borderColor:color,backgroundColor:bgColor,tension:0.2,fill:true,borderWidth:2}]}, options:{responsive:true,maintainAspectRatio:false,scales:{y:{min:1,max:6,title:{display:true,text:'Gj.snitt'}},x:{title:{display:true,text:'Vurderinger'}}},plugins:{tooltip:{mode:'index',intersect:false}}}});
    }

    createLevelGroupsChart() {
        const ctx = document.getElementById('level-groups-chart')?.getContext('2d'); if (!ctx) return;
        const levelData={high:[],medium:[],low:[]}, labels=this.selectedColumnsIndices.map(idx=>this.headers[idx]);
        this.selectedColumnsIndices.forEach(idx => {
            const values = this.statistics[this.headers[idx]]?.values||[];
            levelData.high.push(values.filter(v=>v>=5).length); levelData.medium.push(values.filter(v=>v>=3&&v<5).length); levelData.low.push(values.filter(v=>v<2).length);
        });
        const c={h:this.getCssVar('--color-chart-line-green')||'#4BC0C0', m:this.getCssVar('--color-chart-line-orange')||'#FF9F40', l:this.getCssVar('--color-error')||'#FF6384'};
        const bg={h:this.getCssVarWithAlpha('--color-chart-line-green',0.05), m:this.getCssVarWithAlpha('--color-chart-line-orange',0.05), l:this.getCssVarWithAlpha('--color-error',0.05)};
        if (this.charts.levelGroups) this.charts.levelGroups.destroy();
        this.charts.levelGroups = new Chart(ctx, {type:'line',data:{labels,datasets:[{label:'Høyt (5-6)',data:levelData.high,borderColor:c.h,backgroundColor:bg.h,tension:0.2,fill:'origin',borderWidth:2},{label:'Middels (3-4)',data:levelData.medium,borderColor:c.m,backgroundColor:bg.m,tension:0.2,fill:'origin',borderWidth:2},{label:'Lavt (<2)',data:levelData.low,borderColor:c.l,backgroundColor:bg.l,tension:0.2,fill:'origin',borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true,title:{display:true,text:'Antall elever'}},x:{title:{display:true,text:'Vurderinger'}}},plugins:{tooltip:{mode:'index',intersect:false}}}});
    }

    createHeatmapTable() { 
        const container = document.getElementById('heatmap-chart-table'); if (!container) return; container.innerHTML = ''; 
        if (!this.processedData || !this.processedData.length) { container.innerHTML = '<p>Ingen elevdata.</p>'; return; }
        const table=document.createElement('table'); table.className='heatmap-table'; const thead=table.createTHead(), headerRow=thead.insertRow();
        let th=headerRow.insertCell(); th.textContent='Elev'; th.className='student-name-header';
        this.selectedColumnsIndices.forEach(idx=>{th=headerRow.insertCell();th.textContent=this.headers[idx];th.className='assessment-header';});
        const tbody=table.createTBody();
        this.processedData.forEach(student=>{const row=tbody.insertRow();let cell=row.insertCell();cell.textContent=student.name;cell.className='student-name-cell';
            this.selectedColumnsIndices.forEach(idx=>{const grade=student[this.headers[idx]];cell=row.insertCell();cell.className='grade-cell';
                if(grade!==null&&!isNaN(grade)){cell.textContent=grade.toFixed(1);cell.style.backgroundColor=this.getGradeColor(grade);cell.style.color=this.getTextColorForGrade(grade);}
                else{cell.textContent='–';cell.style.backgroundColor='var(--color-surface)';cell.style.color='var(--color-text-secondary)';}
            });
        });
        container.appendChild(table);
    }

    getGradeColor(grade) { if(grade===null)return 'var(--color-surface)'; if(grade<2)return'#d73027';if(grade<3)return'#fc8d59';if(grade<4)return'#fee08b';if(grade<4.5)return'#ffffbf';if(grade<5)return'#d9ef8b';if(grade<5.5)return'#91cf60';return'#1a9850';}
    getTextColorForGrade(grade) {return(grade>=3&&grade<5)?'var(--color-text)':'#ffffff';}
    
    createDistributionChart() {
        const ctx = document.getElementById('distribution-chart')?.getContext('2d'); if (!ctx) return;
        const counts={'1':0,'2':0,'3':0,'4':0,'5':0,'6':0};
        Object.values(this.statistics).forEach(stat=>{stat?.values?.forEach(g=>{if(g===null||isNaN(g))return; const rg=Math.max(1,Math.min(6,Math.round(g))).toString(); if(counts[rg]!==undefined)counts[rg]++;});});
        if(this.charts.distribution)this.charts.distribution.destroy();
        this.charts.distribution=new Chart(ctx,{type:'bar',data:{labels:Object.keys(counts),datasets:[{label:'Antall',data:Object.values(counts),backgroundColor:[this.getGradeColor(1.5),this.getGradeColor(2.5),this.getGradeColor(3.5),this.getGradeColor(4.5),this.getGradeColor(5.2),this.getGradeColor(5.7)],borderColor:'rgba(var(--color-text-rgb),0.2)',borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,title:{display:true,text:'Antall'}},x:{title:{display:true,text:'Karakter'}}}}});
    }

    createSparklinesWithIndividualDownload() { 
        const container = document.getElementById('student-sparklines'); if (!container) return; container.innerHTML = '';
        if (!this.processedData || !this.processedData.length) { container.innerHTML = '<p>Ingen elever.</p>'; return; }
        this.processedData.forEach(student => { 
            const div = document.createElement('div'); div.className = 'sparkline-item';
            const avgText = student.averageGrade !== null && !isNaN(student.averageGrade) ? student.averageGrade.toFixed(1) : 'N/A';
            const grades = this.selectedColumnsIndices.map(idx => { const g = student[this.headers[idx]]; return (g !== null && !isNaN(g)) ? g : undefined; }); 
            div.innerHTML = `<div class="sparkline-header"><span class="sparkline-name">${student.name}</span><span class="sparkline-average">Snitt: ${avgText}</span><button class="individual-pdf-download-btn" data-student-index="${student.processedDataIndex}" title="PDF">📄</button></div><canvas class="sparkline-chart"></canvas>`;
            container.appendChild(div);
            const canvas = div.querySelector('canvas'); if (!canvas) return;
            new Chart(canvas.getContext('2d'), { type:'line', data:{labels:this.selectedColumnsIndices.map(idx=>this.headers[idx]),datasets:[{data:grades,borderColor:this.getCssVar('--color-primary')||'#21808D',tension:0.4,pointRadius:2.5,fill:false,spanGaps:false,borderWidth:1.5}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:false}},scales:{x:{display:false},y:{display:false,min:0.5,max:6.5}}}});
        });
    }
    
    async generateAllStudentPDFs() {
        this.showProgress(true, 'Forbereder PDFer...');
        const pdfProgressTextEl=document.getElementById('pdf-progress-text'), pdfProgressCountEl=document.getElementById('pdf-progress-count'), pdfGenEl=document.getElementById('pdf-generation-progress');
        if(pdfGenEl)pdfGenEl.classList.remove('hidden'); if(pdfProgressTextEl)pdfProgressTextEl.textContent='Starter...'; if(pdfProgressCountEl)pdfProgressCountEl.textContent='';
        try {
            let jsPDF; if(window.jspdf?.jsPDF)jsPDF=window.jspdf.jsPDF; else if(window.jsPDF)jsPDF=window.jsPDF; else throw new Error('jsPDF mangler.');
            if(!this.processedData||!this.processedData.length){this.showInfo("Ingen data, kan ikke generere PDFer.");if(pdfGenEl)pdfGenEl.classList.add('hidden');this.showProgress(false);return;}
            const blobs=[];
            for(let i=0;i<this.processedData.length;i++){
                const student=this.processedData[i];
                if(pdfProgressTextEl)pdfProgressTextEl.textContent=`Genererer PDF for ${student.name}...`; if(pdfProgressCountEl)pdfProgressCountEl.textContent=`${i+1} av ${this.processedData.length}`;
                await new Promise(r=>setTimeout(r,30)); const blob=await this.generateStudentPDFBlob(student,jsPDF); blobs.push({name:`${student.name.replace(/[^\wæøåÆØÅ\s.-]/g,'_')}_rapport.pdf`,blob});
            }
            if(pdfGenEl)pdfGenEl.classList.add('hidden'); this.showProgress(false); this.offerPDFDownloads(blobs);
        } catch(error){console.error('Feil ved PDF-generering:',error);this.showError(`PDF-feil: ${error.message}`);if(pdfGenEl)pdfGenEl.classList.add('hidden');this.showProgress(false);}
    }
    
    async downloadSingleStudentPDF(student) {
        if(!student){this.showError("Finner ikke elevdata.");return;}
        this.showProgress(true,`Genererer PDF for ${student.name}...`);
        try {
            let jsPDF; if(window.jspdf?.jsPDF)jsPDF=window.jspdf.jsPDF; else if(window.jsPDF)jsPDF=window.jsPDF; else throw new Error('jsPDF mangler.');
            const blob = await this.generateStudentPDFBlob(student,jsPDF); const filename=`${student.name.replace(/[^\wæøåÆØÅ\s.-]/g,'_')}_rapport.pdf`;
            const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download=filename;document.body.appendChild(link);link.click();document.body.removeChild(link);URL.revokeObjectURL(link.href);
            this.showSuccess(`PDF for ${student.name} er generert.`);
        }catch(error){console.error(`Feil for PDF ${student.name}:`,error);this.showError(`Kunne ikke generere PDF: ${error.message}`);}
        finally{this.showProgress(false);}
    }

    offerPDFDownloads(pdfBlobs) {
        const linksList=document.getElementById('pdf-links-list'), container=document.getElementById('pdf-download-links-container');
        if(!linksList||!container){console.error("PDF-lenke container mangler.");return;} linksList.innerHTML='';
        if(!pdfBlobs.length){this.showInfo('Ingen PDFer generert.');container.classList.add('hidden');return;}
        pdfBlobs.forEach(f=>{const l=document.createElement('a');l.href=URL.createObjectURL(f.blob);l.download=f.name;l.className='btn btn--secondary btn--sm';l.style.cssText='display:flex;align-items:center;margin-bottom:var(--space-8);text-align:left;';l.innerHTML=`<span class="message-icon" style="margin-right:8px;font-size:1.1em;">📄</span><span>${f.name}</span>`;linksList.appendChild(l);});
        container.classList.remove('hidden'); this.showSuccess(`${pdfBlobs.length} PDF-rapporter klare.`);
    }

    async generateStudentPDFBlob(student, jsPDFConstructor) {
        const pdf = new jsPDFConstructor({orientation:'portrait',unit:'mm',format:'a4'}); pdf.setFont("Helvetica","normal");
        const pageMargin=15; let yPos=pageMargin; const contentWidth=pdf.internal.pageSize.getWidth()-2*pageMargin; const pageHeight=pdf.internal.pageSize.getHeight(); let currentPage=1;
        const addFooter=()=>{pdf.setFontSize(8);pdf.setTextColor(120,120,120);const d=new Date();pdf.text(`Side ${currentPage} • Rapport: ${d.getDate().toString().padStart(2,'0')}.${(d.getMonth()+1).toString().padStart(2,'0')}.${d.getFullYear()}`,pageMargin,pageHeight-8);}; addFooter();
        if(this.logoDataUrl){try{const logoMaxH=12,logoMaxW=35;const img=new Image();await new Promise((res,rej)=>{img.onload=res;img.onerror=rej;img.src=this.logoDataUrl;});let iW=img.width,iH=img.height;if(iW>0&&iH>0){const aspect=iW/iH;if(iW>logoMaxW){iW=logoMaxW;iH=iW/aspect;}if(iH>logoMaxH){iH=logoMaxH;iW=iH*aspect;}pdf.addImage(this.logoDataUrl,'',pageMargin,yPos,iW,iH);yPos+=iH+3;}else yPos+=3;}catch(e){yPos+=3;}}else yPos+=3;
        pdf.setFontSize(16);pdf.setTextColor(0,0,0);pdf.text('Personlig Elevrapport',pdf.internal.pageSize.getWidth()/2,yPos,{align:'center'});yPos+=10;
        pdf.setFontSize(10.5);pdf.text(`Elev: ${student.name}`,pageMargin,yPos);yPos+=5.5;pdf.text(`Fag: ${document.getElementById('subject').value.trim()||'Ikke oppgitt'}`,pageMargin,yPos);yPos+=5.5;pdf.text(`Skoleår/Periode: ${document.getElementById('school-year').value.trim()||'Ukjent'}`,pageMargin,yPos);yPos+=5.5;pdf.text(`Skole: ${document.getElementById('school-name').value.trim()||'Ukjent'}`,pageMargin,yPos);yPos+=8;
        pdf.setDrawColor(200,200,200);pdf.setLineWidth(0.2);pdf.line(pageMargin,yPos,contentWidth+pageMargin,yPos);yPos+=6;
        pdf.setFontSize(13);pdf.text('Dine karakterer og utvikling',pageMargin,yPos);yPos+=6;pdf.setFontSize(9.5);
        const gradesData=this.selectedColumnsIndices.map(idx=>({name:this.headers[idx],value:student[this.headers[idx]]}));
        const validGrades=gradesData.filter(g=>g.value!==null&&!isNaN(g.value)).map(g=>g.value);
        const avgText=student.averageGrade!==null&&!isNaN(student.averageGrade)?student.averageGrade.toFixed(1):'N/A';
        pdf.text(`Ditt snitt for valgte vurderinger: ${avgText}`,pageMargin,yPos);yPos+=6;
        gradesData.forEach(item=>{if(yPos>pageHeight-pageMargin-20){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}const gradeTxt=item.value!==null&&!isNaN(item.value)?item.value.toFixed(1):'Ikke vurdert';pdf.text(`${item.name}: ${gradeTxt}`,pageMargin+5,yPos);yPos+=5;});yPos+=4;
        if(yPos>pageHeight-pageMargin-30){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}pdf.setFontSize(13);pdf.text('Din personlige utvikling',pageMargin,yPos);yPos+=6;pdf.setFontSize(9.5);
        if(validGrades.length>=2){const first=validGrades[0],last=validGrades[validGrades.length-1],imp=parseFloat((last-first).toFixed(1));let devTxt="";if(imp>0.1)devTxt=`Positiv utvikling! Karakteren økte med ${imp} poeng, fra ${first.toFixed(1)} til ${last.toFixed(1)}.`;else if(imp<-0.1)devTxt=`Karakteren gikk ned ${Math.abs(imp)} poeng, fra ${first.toFixed(1)} til ${last.toFixed(1)}. Reflekter over strategier.`;else devTxt=`Stabilt nivå (${first.toFixed(1)}) mellom første og siste vurdering. God konsistens.`;const splitT=pdf.splitTextToSize(devTxt,contentWidth);pdf.text(splitT,pageMargin,yPos);yPos+=(splitT.length*4.5)+3;}
        else if(validGrades.length===1){if(yPos>pageHeight-pageMargin-20){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}pdf.text(`Én registrert karakter (${validGrades[0].toFixed(1)}). Trenger flere for å se utvikling.`,pageMargin,yPos,{maxWidth:contentWidth});yPos+=9;}
        else{if(yPos>pageHeight-pageMargin-20){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}pdf.text('Ingen karakterer registrert i valgte vurderinger. Utvikling kan ikke vises.',pageMargin,yPos,{maxWidth:contentWidth});yPos+=9;}
        if(yPos>pageHeight-pageMargin-30){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}
        if(validGrades.length>0){const h=Math.max(...validGrades),l=Math.min(...validGrades);pdf.text(`Høyeste karakter: ${h.toFixed(1)}`,pageMargin,yPos);yPos+=5;pdf.text(`Laveste karakter: ${l.toFixed(1)}`,pageMargin,yPos);yPos+=5;if(h-l>0.1){pdf.text(`Spredning: ${(h-l).toFixed(1)} poeng.`,pageMargin,yPos,{maxWidth:contentWidth});yPos+=5;}}yPos+=4;
        const goalsY=yPos;if(goalsY>pageHeight-pageMargin-50&&gradesData.length>1){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}else if(goalsY>pageHeight-pageMargin-20){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}
        pdf.setFontSize(13);pdf.text('Refleksjon og mål',pageMargin,yPos);yPos+=6;pdf.setFontSize(9.5);
        const goals=['• Hva er du mest fornøyd med?','• Hvilke strategier fungerte bra?','• Temaer/ferdigheter å jobbe mer med?','• Sett 1-2 konkrete læringsmål.','• Diskuter med læreren din.'];
        goals.forEach(g=>{if(yPos>pageHeight-pageMargin-15){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}const splitG=pdf.splitTextToSize(g,contentWidth-2);pdf.text(splitG,pageMargin,yPos);yPos+=(splitG.length*4.5)+1.5;});yPos+=4;
        const chartY=yPos,estChartH=70;if(yPos+estChartH>pageHeight-pageMargin-10){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}else yPos=chartY;
        if(validGrades.length>0)await this.addStudentChartToPDF(pdf,student,yPos,pageMargin);else{if(yPos>pageHeight-pageMargin-15){pdf.addPage();currentPage++;addFooter();yPos=pageMargin;}pdf.text("Graf kan ikke vises (mangler data).",pageMargin,yPos,{maxWidth:contentWidth});}
        return pdf.output('blob');
    }

    async addStudentChartToPDF(pdf, student, startY, chartPageMargin) { 
        const tempCanvas = document.createElement('canvas'), dpiScale = 2.5, chartContentWidthMM = pdf.internal.pageSize.getWidth()-2*chartPageMargin;
        const chartWidthPx = Math.round(chartContentWidthMM*(96/25.4)*(dpiScale/2.2)), chartHeightPx = Math.round(65*(96/25.4)*(dpiScale/2.2));
        tempCanvas.width = chartWidthPx; tempCanvas.height = chartHeightPx;
        const ctx = tempCanvas.getContext('2d'); if (!ctx) return;
        const grades = this.selectedColumnsIndices.map(idx=>{const g=student[this.headers[idx]];return(g!==null&&!isNaN(g))?g:undefined;});
        const labels = this.selectedColumnsIndices.map(idx=>this.headers[idx]);
        const chart = new Chart(ctx, {type:'line',data:{labels,datasets:[{label:'Dine karakterer',data:grades,borderColor:this.getCssVar('--color-chart-line-blue')||'#0D6EFD',backgroundColor:this.getCssVarWithAlpha('--color-chart-line-blue',0.05),fill:true,tension:0.2,pointRadius:2.2*dpiScale,pointHoverRadius:3.5*dpiScale,borderWidth:1.5*dpiScale,spanGaps:false}]},options:{responsive:false,animation:false,devicePixelRatio:dpiScale,plugins:{legend:{display:true,position:'bottom',labels:{font:{size:7*dpiScale},padding:4*dpiScale}},title:{display:true,text:'Din karakterutvikling',font:{size:9*dpiScale,weight:'500'},padding:{top:4*dpiScale,bottom:6*dpiScale}}},scales:{y:{min:0.5,max:6.5,title:{display:true,text:'Karakter',font:{size:7*dpiScale}},ticks:{font:{size:6*dpiScale},stepSize:1,padding:2*dpiScale}},x:{title:{display:true,text:'Vurderinger',font:{size:7*dpiScale}},ticks:{font:{size:6*dpiScale},autoSkip:true,maxRotation:45,minRotation:20,padding:2*dpiScale}}},layout:{padding:{top:4*dpiScale,right:6*dpiScale,bottom:4*dpiScale,left:4*dpiScale}}}});
        await new Promise(resolve => setTimeout(() => {
            try {
                const imgData = tempCanvas.toDataURL('image/png',0.95);
                const finalImgWidth=chartContentWidthMM, finalImgHeight=(tempCanvas.height/tempCanvas.width)*finalImgWidth;
                pdf.addImage(imgData,'PNG',chartPageMargin,startY,finalImgWidth,finalImgHeight);
            } catch(e){console.error("Feil ved canvas til DataURL:",e);}
            finally{chart.destroy();resolve();}
        },350)); 
    }

    showProgress(show, actionText='Laster...'){const upA=document.getElementById('upload-area'),prA=document.getElementById('upload-progress'),prTxtEl=document.getElementById('progress-action-text');if(upA&&prA){if(show){upA.classList.add('hidden');prA.classList.remove('hidden');if(prTxtEl)prTxtEl.textContent=actionText;}else{upA.classList.remove('hidden');prA.classList.add('hidden');}}}
    updateProgress(percentage,text){const f=document.getElementById('progress-fill'),tE=document.getElementById('progress-text');if(f)f.style.width=`${Math.min(100,Math.max(0,percentage))}%`;if(tE)tE.textContent=text;}
    showConfigSection(){const cS=document.getElementById('config-section');if(cS)cS.classList.remove('hidden');}
    showError(message){this.showMessage(message,'error-message');}
    showSuccess(message){this.showMessage(message,'success-message');}
    showInfo(message){this.showMessage(message,'info-message');}
    showMessage(message,elementId){const msgEl=document.getElementById(elementId);if(!msgEl){alert(message);return;}const txtC=msgEl.querySelector('.message-text');if(txtC)txtC.textContent=message;else msgEl.textContent=message;['error-message','success-message','info-message'].forEach(id=>{const el=document.getElementById(id);if(el&&id!==elementId)el.classList.add('hidden');});msgEl.classList.remove('hidden');setTimeout(()=>{if(msgEl)msgEl.classList.add('hidden');},7000);}
}

document.addEventListener('DOMContentLoaded',()=>{const core=['file-input','upload-area','logo-input','plus-minus-handling','analyze-btn','generate-pdfs-btn','config-section','results-section'];let ok=true;core.forEach(id=>{if(!document.getElementById(id)){console.error(`HTML-element '${id}' mangler!`);ok=false;}});if(ok)new ExcelStatsApp();else alert("Kjerne-HTML mangler. Sjekk konsoll.");});

