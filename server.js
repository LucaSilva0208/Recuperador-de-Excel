const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const crypto = require('crypto');




class FileManager {
    constructor(uploadDir, recoveryDir, retentionTimeMs = 600000) {
        this.uploadDir = uploadDir;
        this.recoveryDir = recoveryDir;
        this.retentionTimeMs = retentionTimeMs;

        
       
        if (!fs.existsSync(this.uploadDir)) fs.mkdirSync(this.uploadDir);
        if (!fs.existsSync(this.recoveryDir)) fs.mkdirSync(this.recoveryDir);
    }

    
  
    cleanOldFiles() {
        const now = Date.now();
        [this.uploadDir, this.recoveryDir].forEach(dir => {
            fs.readdir(dir, (err, files) => {
                if (err) return console.error(`Erro ao ler diretório ${dir}:`, err);
                
                files.forEach(file => {
                    const filePath = path.join(dir, file);
                    fs.stat(filePath, (err, stats) => {
                        if (err) return;
                        
                        
                        if (now - stats.birthtimeMs > this.retentionTimeMs) {
                            fs.unlink(filePath, () => console.log(`[Auto-Clean] Removido: ${file}`));
                        }
                    });
                });
            });
        });
    }

    startAutoCleanup(intervalMs = 300000) {
        
        setInterval(() => this.cleanOldFiles(), intervalMs);
        console.log("Agendador de limpeza de arquivos iniciado.");
    }
}


class ExcelRecoveryService {
    constructor() {
    }

    _sanitizeCell(cell) {
        if (cell === undefined || cell === null) return cell;
        
        
        if (typeof cell === 'object' && !(cell instanceof Date)) {
            if (cell.result !== undefined) cell = cell.result; 
            else if (cell.richText) cell = cell.richText.map(rt => rt.text).join('');
            else if (cell.text) cell = cell.text;
            else cell = '[DATA_CORRUPTED]';
        }
        if (typeof cell === 'string') {
            
            return cell.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
        }
        return cell;
    }
    async recoverFile(inputPath, outputPath, password, duplicateAction = 'ask') {
        try {
            console.log(`Lendo arquivo em formato Stream via ExcelJS...`);
            
            const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(inputPath, {
                worksheets: 'emit',
                sharedStrings: 'cache',
                hyperlinks: 'cache',
                styles: 'cache',
                emptyStrings: true
            });

            const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter({ 
                filename: outputPath,
                useStyles: true
            });
            
            let stats = { sheets: 0, totalRows: 0, truncated: false };

            for await (const worksheetReader of workbookReader) {
                stats.sheets++;
                console.log(`Processando aba: ${worksheetReader.name}`);
                
                const seenRows = new Set();
                let targetRowNumber = 1;
                
                
                const sheetOptions = {};
                if (worksheetReader.properties) sheetOptions.properties = worksheetReader.properties;
                if (worksheetReader.pageSetup) sheetOptions.pageSetup = worksheetReader.pageSetup;
                if (worksheetReader.views) sheetOptions.views = worksheetReader.views;
                if (worksheetReader.state) sheetOptions.state = worksheetReader.state;

                const worksheetWriter = workbookWriter.addWorksheet(worksheetReader.name, sheetOptions);
                
                
                if (worksheetReader.columns) {
                    worksheetReader.columns.forEach((col, index) => {
                        if (col) {
                            const newCol = worksheetWriter.getColumn(index + 1);
                            if (col.width) newCol.width = col.width;
                            if (col.hidden) newCol.hidden = col.hidden;
                            if (col.style) newCol.style = col.style;
                            if (col.outlineLevel) newCol.outlineLevel = col.outlineLevel;
                        }
                    });
                }
                
                for await (const row of worksheetReader) {
                    let rowValues = [];
                    let hasData = false;

                    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        const val = this._sanitizeCell(cell.value);
                        rowValues[colNumber] = val;
                        if (val !== null && val !== undefined && val !== '') hasData = true;
                    });
                    
                    if (hasData) {
                        const rowString = JSON.stringify(rowValues);
                        const rowHash = crypto.createHash('md5').update(rowString).digest('hex');
                        
                        if (seenRows.has(rowHash)) {
                            if (duplicateAction === 'ask') {
                                throw new Error("DUPLICATES_FOUND");
                            } else if (duplicateAction === 'remove') {
                                continue; // Pula a escrita desta linha (remove a duplicata)
                            }
                        }
                        seenRows.add(rowHash);
                    }
                    
                    const newRow = worksheetWriter.getRow(duplicateAction === 'remove' ? targetRowNumber : row.number);
                    
                    
                    if (row.height) newRow.height = row.height;
                    if (row.hidden) newRow.hidden = row.hidden;
                    if (row.style) newRow.style = row.style;

                    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        const newCell = newRow.getCell(colNumber);
                        newCell.value = this._sanitizeCell(cell.value);
                        if (cell.style) newCell.style = cell.style;
                    });
                    
                    newRow.commit();
                    stats.totalRows++;
                    targetRowNumber++;
                }
                
                
                const merges = worksheetReader.merges || (worksheetReader.model && worksheetReader.model.merges);
                if (merges && Array.isArray(merges)) {
                    merges.forEach(merge => {
                        try { worksheetWriter.mergeCells(merge); } catch (e) {  }
                    });
                }

                worksheetWriter.commit();
            }
            
            if (stats.sheets === 0) {
                throw new Error("Nenhuma aba válida pôde ser lida. Certifique-se que o arquivo é um .xlsx válido.");
            }

            console.log("Salvando arquivo final...");
            await workbookWriter.commit();

            return { outputPath, stats };
        } catch (error) {
            throw new Error(`Falha fatal na leitura do arquivo. Detalhes: ${error.message}`);
        }
    }
}


class RecoveryController {
    constructor(recoveryService, fileManager) {
        this.service = recoveryService;
        this.fileManager = fileManager;
    }

    async handleRecovery(req, res, next) {
        if (!req.file) return res.status(400).json({ error: 'Arquivo inválido ou inexistente.' });

        const inputPath = req.file.path;
        const outputFilename = `recovered_${path.basename(req.file.path)}`;
        const outputPath = path.join(this.fileManager.recoveryDir, outputFilename);
        const password = req.body.password; // Captura a senha do formulário
        const duplicateAction = req.body.duplicateAction || 'ask'; // 'ask', 'keep', 'remove'

        console.log(`Processando: ${req.file.originalname} | DuplicateAction: ${duplicateAction}`);

        try {
            const result = await this.service.recoverFile(inputPath, outputPath, password, duplicateAction);
            
           
            res.set('X-Recovery-Sheets', result.stats.sheets);
            res.set('X-Recovery-Rows', result.stats.totalRows);
            res.set('X-Recovery-Truncated', result.stats.truncated ? 'true' : 'false');
            
            
            res.set('Access-Control-Expose-Headers', 'X-Recovery-Sheets, X-Recovery-Rows, X-Recovery-Truncated, Content-Disposition');

            res.download(result.outputPath, `recovered_${req.file.originalname}`, (err) => {
                if (err) console.error("Erro no download:", err);
                
                
                fs.unlink(result.outputPath, () => {});
            });
        } catch (error) {
            if (error.message === "DUPLICATES_FOUND") {
                return res.status(409).json({ 
                    error: 'DUPLICATES_FOUND', 
                    message: 'O arquivo contém informações duplicadas. É necessário manter as cópias?' 
                });
            }
            if (error.message.includes("senha")) {
                return res.status(422).json({ error: error.message });
            }
            next(error);
        } finally {
            
            fs.unlink(inputPath, () => {});
        }
    }
}


const app = express();
const UPLOAD_DIR = path.join(__dirname, 'temp_uploads');
const RECOVERY_DIR = path.join(__dirname, 'temp_recovered');


const fileManager = new FileManager(UPLOAD_DIR, RECOVERY_DIR, 300000);
const recoveryService = new ExcelRecoveryService();
const recoveryController = new RecoveryController(recoveryService, fileManager);

fileManager.startAutoCleanup();


const upload = multer({ 
    storage: multer.diskStorage({
        destination: (req, file, cb) => cb(null, UPLOAD_DIR),
        filename: (req, file, cb) => cb(null, `${uuidv4()}_${file.originalname}`)
    }),
    limits: { fileSize: 25 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        const allowedMimes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel',
            'application/octet-stream'
        ];
        if (allowedMimes.includes(file.mimetype) || file.originalname.match(/\.(xlsx)$/i)) {
            cb(null, true);
        } else {
            cb(new Error('Apenas arquivos Excel modernos (.xlsx) são permitidos. Formato .xls antigo não é suportado pelo motor de Streams.'));
        }
    }
});


app.post('/recover', upload.single('file'), (req, res, next) => recoveryController.handleRecovery(req, res, next));


app.use(express.static(path.join(__dirname, 'public')));


app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});


app.use((err, req, res, next) => {
    if (err instanceof multer.MulterError) return res.status(400).json({ error: `Erro de Upload: ${err.message}` });
    if (err.message.includes('Apenas arquivos Excel')) return res.status(400).json({ error: err.message });
    console.error(err);
    res.status(500).json({ error: "Erro interno no servidor." });
});

const PORT = process.env.PORT || 3000;
const server = app.listen(PORT, () => {
    console.log(`Servidor de Recuperação Excel rodando na porta ${PORT}`);
});


server.setTimeout(120000);
