const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');


const API_URL = 'http://localhost:3000/recover';
const TEST_FILE = path.join(__dirname, 'teste_origem.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'teste_recuperado.xlsx');

async function createTestFile() {
    console.log("Criando arquivo de teste com duplicidade proposital...");
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
        ["ID", "Nome", "Status", "Obs"],
        [1, "Teste A", "Ativo", "Dados normais"],
        [2, "Teste B", "Erro", "Caractere\x00Inválido"],
        [1, "Teste A", "Ativo", "Dados normais"], // Linha Duplicada Injetada!
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Dados");
    XLSX.writeFile(wb, TEST_FILE);
}

async function uploadTest(duplicateAction = null) {
    const { fetch, FormData } = globalThis;
    const fileBlob = new Blob([fs.readFileSync(TEST_FILE)]);
    
    const formData = new FormData();
    formData.append('file', fileBlob, 'teste_origem.xlsx');
    if (duplicateAction) {
        formData.append('duplicateAction', duplicateAction);
    }

    console.log(`\nEnviando para ${API_URL} (Ação: ${duplicateAction || 'ask'})...`);
    const startTime = Date.now();
    
    const response = await fetch(API_URL, {
        method: 'POST',
        body: formData
    });

    if (!response.ok) {
        const errorData = await response.json().catch(() => null);
        if (response.status === 409 && errorData && errorData.error === 'DUPLICATES_FOUND') {
            console.log("⚠️ Servidor encontrou duplicidades e barrou com sucesso!");
            return 'DUPLICATES_FOUND';
        }
        throw new Error(`Erro do Servidor (${response.status})`);
    }

    const buffer = await response.arrayBuffer();
    const outName = OUTPUT_FILE.replace('.xlsx', `_${duplicateAction}.xlsx`);
    fs.writeFileSync(outName, Buffer.from(buffer));
    
    const duration = Date.now() - startTime;
    const rows = response.headers.get('X-Recovery-Rows');
    const sheets = response.headers.get('X-Recovery-Sheets');

    console.log(`✅ SUCESSO! Recuperado em ${duration}ms`);
    console.log(`💾 Salvo em: ${outName} | Total de Linhas Resultantes: ${rows}`);
    return 'SUCCESS';
}

async function runTest() {
    console.log("=== Iniciando Bateria de Testes ===");
    if (!fs.existsSync(TEST_FILE)) await createTestFile();

    try {
        // 1. Testa bloqueio de duplicidade (Sem ação definida)
        const checkStatus = await uploadTest();
        
        if (checkStatus === 'DUPLICATES_FOUND') {
            // 2. Testa usuário pedindo para Remover Duplicidades
            await uploadTest('remove');
            
            // 3. Testa usuário pedindo para Manter as Cópias (ignorar verificação)
            await uploadTest('keep');
        }
    } catch (error) {
        console.error("\n❌ FALHA NO TESTE:", error.message);
    }
}

runTest();
