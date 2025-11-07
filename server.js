const http = require('http');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Nomes dos arquivos e planilhas (ajuste se necessário)
const ARQUIVO_RELATORIO = 'relatorio.xlsx';
const ARQUIVO_ELEGIVEIS = 'elegiveis_auto.xlsx';
const ARQUIVO_BITRIX = 'baixada_do_bitrix.xlsx';
const ARQUIVO_TIME = 'time_novembro.xlsx';

const NOME_PLANILHA_PRINCIPAL = 'Sheet1'; // Nome da planilha principal em 'elegiveis_auto'
const NOME_PLANILHA_RELACIONAMENTO = 'C6 - Relacionamento';
const NOME_PLANILHA_SUPERVISORES = 'supervisores';

// Função para registrar logs
const log = (message) => {
    console.log(message);
    // Em um cenário real, você enviaria isso para o front-end via WebSockets
};

// Função principal da automação
function runAutomation() {
    try {
        log('Iniciando automação...');

        // --- Passo 1 e 2: Copiar relatório e adicionar colunas ---
        log(`Lendo ${ARQUIVO_RELATORIO}...`);
        const relatorioWb = xlsx.readFile(ARQUIVO_RELATORIO);
        const relatorioWs = relatorioWb.Sheets[relatorioWb.SheetNames[0]];
        const dataRelatorio = xlsx.utils.sheet_to_json(relatorioWs, { header: 1, defval: null });

        log('Adicionando colunas em branco em C, D, E, F...');
        const dataComNovasColunas = dataRelatorio.map(row => {
            const newRow = [...row];
            newRow.splice(2, 0, null, null, null, null); // Insere 4 valores nulos a partir do índice 2 (Coluna C)
            return newRow;
        });

        const elegiveisAutoWb = xlsx.utils.book_new();
        const elegiveisAutoWs = xlsx.utils.aoa_to_sheet(dataComNovasColunas, { cellDates: true });
        xlsx.utils.book_append_sheet(elegiveisAutoWb, elegiveisAutoWs, NOME_PLANILHA_PRINCIPAL);
        log('Planilha principal de "elegiveis_auto" criada na memória.');

        // --- Passo 3, 4, 5, 6: Processar planilhas de apoio ---
        log(`Lendo ${ARQUIVO_BITRIX}...`);
        const bitrixWb = xlsx.readFile(ARQUIVO_BITRIX);
        const bitrixWs = bitrixWb.Sheets[bitrixWb.SheetNames[0]];
        const bitrixData = xlsx.utils.sheet_to_json(bitrixWs, { header: 'A' });
        const bitrixColunas = bitrixData.map(row => [row['CNPJ'], row['FASE'], row['RESPONSAVEL']]);
        const relacionamentoWs = xlsx.utils.aoa_to_sheet(bitrixColunas);
        xlsx.utils.book_append_sheet(elegiveisAutoWb, relacionamentoWs, NOME_PLANILHA_RELACIONAMENTO);
        log(`Planilha "${NOME_PLANILHA_RELACIONAMENTO}" adicionada.`);

        log(`Lendo ${ARQUIVO_TIME}...`);
        const timeWb = xlsx.readFile(ARQUIVO_TIME);
        const timeWs = timeWb.Sheets[timeWb.SheetNames[0]];
        const timeData = xlsx.utils.sheet_to_json(timeWs, { header: 'A' });
        const timeColunas = timeData.map(row => [row['Consultor'], row['Equipe']]);
        const supervisoresWs = xlsx.utils.aoa_to_sheet(timeColunas);
        xlsx.utils.book_append_sheet(elegiveisAutoWb, supervisoresWs, NOME_PLANILHA_SUPERVISORES);
        log(`Planilha "${NOME_PLANILHA_SUPERVISORES}" adicionada.`);

        // --- Passo 7: Fórmulas PROCV (VLOOKUP) ---
        log('Aplicando fórmulas PROCV...');
        const mainSheetData = xlsx.utils.sheet_to_json(elegiveisAutoWs, { header: 1 });
        
        // Renomeia cabeçalhos
        mainSheetData[0][2] = 'fase';
        mainSheetData[0][3] = 'responsável';
        mainSheetData[0][4] = 'Supervisor';
        mainSheetData[0][5] = 'Coluna F'; // Nome para a coluna F

        for (let i = 1; i < mainSheetData.length; i++) { // Começa de 1 para pular o cabeçalho
            const rowIndex = i + 1;
            // PROCV para Fase (Coluna C)
            mainSheetData[i][2] = { f: `VLOOKUP(B${rowIndex},'${NOME_PLANILHA_RELACIONAMENTO}'!A:B,2,FALSE)` };
            // PROCV para Responsável (Coluna D)
            mainSheetData[i][3] = { f: `VLOOKUP(B${rowIndex},'${NOME_PLANILHA_RELACIONAMENTO}'!A:C,3,FALSE)` };
            // PROCV para Supervisor (Coluna E)
            mainSheetData[i][4] = { f: `VLOOKUP(D${rowIndex},'${NOME_PLANILHA_SUPERVISORES}'!A:B,2,FALSE)` };
        }

        const wsComFormulas = xlsx.utils.aoa_to_sheet(mainSheetData);
        elegiveisAutoWb.Sheets[NOME_PLANILHA_PRINCIPAL] = wsComFormulas;
        
        // Salva temporariamente para que o Excel calcule as fórmulas
        xlsx.writeFile(elegiveisAutoWb, 'temp_calculo.xlsx');
        const wbCalculado = xlsx.readFile('temp_calculo.xlsx', { cellFormula: false }); // Lê os valores, não as fórmulas
        log('Fórmulas calculadas e convertidas para valores.');

        // --- Passo 8: Filtros e exclusão de linhas ---
        log('Aplicando filtros e removendo linhas...');
        const wsFinal = wbCalculado.Sheets[NOME_PLANILHA_PRINCIPAL];
        let finalData = xlsx.utils.sheet_to_json(wsFinal, { defval: null });

        // Encontrar os nomes exatos das colunas de filtro
        const headers = Object.keys(finalData[0]);
        const colElegivel = headers.find(h => h.toUpperCase().includes('FL_ELEGIVEL_VENDA_C6PAY'));
        const colTipoPessoa = headers.find(h => h.toUpperCase().includes('TIPO_PESSOA'));
        const colDataAprovacao = headers.find(h => h.toUpperCase().includes('DATA_APROVAÇÃO_PAY'));
        const colStatusCC = headers.find(h => h.toUpperCase().includes('STATUS_CC'));

        const dadosFiltrados = finalData.filter(row => {
            const isElegivel = row[colElegivel] === 1 || row[colElegivel] === '1';
            const isPj = row[colTipoPessoa] === 'PJ';
            const isDataVazia = row[colDataAprovacao] === null || row[colDataAprovacao] === '';
            const isStatusLiberada = row[colStatusCC] === 'Liberada';
            
            return isElegivel && isPj && isDataVazia && isStatusLiberada;
        });
        log(`${dadosFiltrados.length} linhas restantes após o filtro.`);

        // --- Final: Salvar o arquivo final ---
        const finalWsComFiltro = xlsx.utils.json_to_sheet(dadosFiltrados);
        wbCalculado.Sheets[NOME_PLANILHA_PRINCIPAL] = finalWsComFiltro;
        
        xlsx.writeFile(wbCalculado, ARQUIVO_ELEGIVEIS);
        fs.unlinkSync('temp_calculo.xlsx'); // Remove o arquivo temporário
        log(`Processo concluído! Arquivo "${ARQUIVO_ELEGIVEIS}" salvo com sucesso.`);

    } catch (error) {
        log(`Ocorreu um erro: ${error.message}`);
        console.error(error);
    }
}

// Servidor HTTP simples para servir a página HTML
const server = http.createServer((req, res) => {
    if (req.method === 'POST' && req.url === '/run') {
        runAutomation();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ message: 'Processo iniciado. Verifique o console do servidor para logs.' }));
    } else {
        fs.readFile(path.join(__dirname, 'index.html'), (err, data) => {
            if (err) {
                res.writeHead(500);
                res.end('Erro ao carregar index.html');
                return;
            }
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(data);
        });
    }
});

const PORT = 3000;
server.listen(PORT, () => {
    console.log(`Servidor rodando em http://localhost:${PORT}`);
    console.log('Coloque os arquivos .xlsx na mesma pasta deste script.');
    console.log('Abra o navegador e acesse a URL acima para iniciar.');
});
