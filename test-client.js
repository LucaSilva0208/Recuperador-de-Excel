const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');


const API_URL = 'http://localhost:3000/recover';
const TEST_FILE = path.join(__dirname, 'teste_origem.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'teste_recuperado.xlsx');

async function runTest() {
    console.log("=== Iniciando Teste do Servidor ===");

    
    if (!fs.existsSync(TEST_FILE)) {
        console.log("Criando arquivo de teste...");
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([
            ["ID", "Nome", "Status", "Obs"],
            [1, "Teste A", "Ativo", "Dados normais"],
            [2, "Teste B", "Erro", "Caractere\x00Inválido"],
        ]);
        XLSX.utils.book_append_sheet(wb, ws, "Dados");
        XLSX.writeFile(wb, TEST_FILE);
    }

    
    const { fetch, FormData } = globalThis;
    const fileBlob = new Blob([fs.readFileSync(TEST_FILE)]);
    
    const formData = new FormData();
    formData.append('file', fileBlob, 'teste_origem.xlsx');

    try {
        console.log(`Enviando para ${API_URL}...`);
        const startTime = Date.now();
        
        const response = await fetch(API_URL, {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Erro do Servidor (${response.status}): ${errorText}`);
        }

        
        const buffer = await response.arrayBuffer();
        fs.writeFileSync(OUTPUT_FILE, Buffer.from(buffer));
        
        const duration = Date.now() - startTime;
        
        
        const rows = response.headers.get('X-Recovery-Rows');
        const sheets = response.headers.get('X-Recovery-Sheets');

        console.log(`\n✅ SUCESSO! Recuperado em ${duration}ms`);
        console.log(`💾 Salvo em: ${OUTPUT_FILE}`);

    } catch (error) {
        console.error("\n❌ FALHA NO TESTE:", error.message);
    }
}

runTest();
