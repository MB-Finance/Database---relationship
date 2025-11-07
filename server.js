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

// Função para verificar se os arquivos necessários existem
function checkRequiredFiles() {
    const files = [ARQUIVO_RELATORIO, ARQUIVO_BITRIX, ARQUIVO_TIME];
    for (const file of files) {
        if (!fs.existsSync(file)) {
            throw new Error(`Arquivo necessário não encontrado: ${file}`);
        }
    }
    log('Todos os arquivos de entrada foram encontrados.');
    return true;
}

// Função principal da automação
function runAutomation() {
    try {
        log('Iniciando automação...');
        // --- Passo 0: Verificar arquivos ---
        checkRequiredFiles();

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
        const bitrixData = xlsx.utils.sheet_to_json(bitrixWs); // Lê usando a primeira linha como cabeçalho

        // Encontra os nomes reais dos cabeçalhos
        const bitrixHeaders = Object.keys(bitrixData[0] || {});
        const hCnpj = bitrixHeaders.find(h => h.toUpperCase().includes('CNPJ'));
        const hFase = bitrixHeaders.find(h => h.toUpperCase().includes('Fase'));
        const hResponsavel = bitrixHeaders.find(h => h.toUpperCase().includes('Responsável'));

        const bitrixColunas = bitrixData.map(row => [row[hCnpj], row[hFase], row[hResponsavel]]);
        const relacionamentoWs = xlsx.utils.aoa_to_sheet(bitrixColunas);
        xlsx.utils.book_append_sheet(elegiveisAutoWb, relacionamentoWs, NOME_PLANILHA_RELACIONAMENTO);
        log(`Planilha "${NOME_PLANILHA_RELACIONAMENTO}" adicionada.`);

        log(`Lendo ${ARQUIVO_TIME}...`);
        const timeWb = xlsx.readFile(ARQUIVO_TIME);
        const timeWs = timeWb.Sheets[timeWb.SheetNames[0]];
        const timeData = xlsx.utils.sheet_to_json(timeWs); // Lê usando a primeira linha como cabeçalho
        
        const timeHeaders = Object.keys(timeData[0] || {});
        const hConsultor = timeHeaders.find(h => h.toUpperCase().includes('Consultor'));
        const hEquipe = timeHeaders.find(h => h.toUpperCase().includes('EQUIPE'));

        const timeColunas = timeData.map(row => [row[hConsultor], row[hEquipe]]);
        const supervisoresWs = xlsx.utils.aoa_to_sheet(timeColunas);
        xlsx.utils.book_append_sheet(elegiveisAutoWb, supervisoresWs, NOME_PLANILHA_SUPERVISORES);
        log(`Planilha "${NOME_PLANILHA_SUPERVISORES}" adicionada.`);

        // --- Passo 7: Fórmulas PROCV (VLOOKUP) ---
        log('Aplicando fórmulas PROCV...');
        const mainSheetData = xlsx.utils.sheet_to_json(elegiveisAutoWs, { header: 1 });
        
        // Define os cabeçalhos APENAS para as novas colunas inseridas
        if (mainSheetData.length > 0) {
            mainSheetData[0][2] = 'fase';
            mainSheetData[0][3] = 'responsável';
            mainSheetData[0][4] = 'Supervisor';
            // A coluna F (índice 5) não foi solicitada para renomear, então a deixamos como está.
        }

        for (let i = 1; i < mainSheetData.length; i++) { // Começa de 1 para pular o cabeçalho
            const rowIndex = i + 1;
            // PROCV para Fase (Coluna C)
            mainSheetData[i][2] = { f: `XLOOKUP(B${rowIndex},'${NOME_PLANILHA_RELACIONAMENTO}'!A:A,'${NOME_PLANILHA_RELACIONAMENTO}'!B:B)` };
            // PROCV para Responsável (Coluna D)
            mainSheetData[i][3] = { f: `XLOOKUP(B${rowIndex},'${NOME_PLANILHA_RELACIONAMENTO}'!A:A,'${NOME_PLANILHA_RELACIONAMENTO}'!C:C)` };
            // PROCV para Supervisor (Coluna E)
            mainSheetData[i][4] = { f: `XLOOKUP(D${rowIndex},'${NOME_PLANILHA_SUPERVISORES}'!A:A,'${NOME_PLANILHA_SUPERVISORES}'!B:B)` };
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
        
        if (finalData.length === 0) {
            log('AVISO: A planilha principal está vazia após o cálculo das fórmulas. O arquivo final estará vazio.');
            xlsx.writeFile(wbCalculado, ARQUIVO_ELEGIVEIS);
            fs.unlinkSync('temp_calculo.xlsx');
            return;
        }
 
        // Função auxiliar para encontrar cabeçalhos de forma flexível
        const findHeader = (headers, primaryName, fallbackColumnLetter) => {
            // 1. Tenta encontrar pelo nome primário (ignorando maiúsculas/minúsculas e espaços)
            let header = headers.find(h => h.trim().toUpperCase() === primaryName.toUpperCase());
            if (header) return header;

            // 2. Se não encontrar, usa a letra da coluna como fallback
            const colIndex = xlsx.utils.decode_col(fallbackColumnLetter);
            if (headers[colIndex]) {
                log(`AVISO: Coluna "${primaryName}" não encontrada pelo nome. Usando a coluna de fallback '${fallbackColumnLetter}' (${headers[colIndex]}).`);
                return headers[colIndex];
            }

            return null; // Não encontrou de nenhuma forma
        };

        const headers = Object.keys(finalData[0]);
        const colElegivel = findHeader(headers, 'FL_ELEGIVEL_VENDA_C6PAY', 'AK');
        const colTipoPessoa = findHeader(headers, 'TIPO_PESSOA', 'H');
        const colDataAprovacao = findHeader(headers, 'DT_APROVACAO_PAY', 'AM'); // Usando AK como fallback, conforme prompt anterior
        const colStatusCC = findHeader(headers, 'STATUS_CC', 'Y');
        
        if (!colElegivel || !colTipoPessoa || !colDataAprovacao || !colStatusCC) {
            // O log de erro agora será mais específico sobre qual coluna falhou
            throw new Error(`Não foi possível encontrar uma ou mais colunas de filtro, mesmo com as coordenadas de fallback. Verifique os nomes e as posições das colunas na planilha de relatório.`);
        }

        log(`Total de linhas antes do filtro: ${finalData.length}`);

        const filtro1 = finalData.filter(row => row[colElegivel] === 1 || row[colElegivel] === '1');
        log(`Linhas restantes após filtro FL_ELEGIVEL_VENDA_C6PAY = "1": ${filtro1.length}`);

        const filtro2 = filtro1.filter(row => row[colTipoPessoa] === 'PJ');
        log(`Linhas restantes após filtro Tipo_Pessoa = "PJ": ${filtro2.length}`);

        const filtro3 = filtro2.filter(row => row[colDataAprovacao] === null || row[colDataAprovacao] === '');
        log(`Linhas restantes após filtro data_aprovação_pay = vazias: ${filtro3.length}`);

        const dadosFiltrados = filtro3.filter(row => row[colStatusCC] === 'LIBERADA');
        log(`Linhas restantes após filtro STATUS_CC = "LIBERADA": ${dadosFiltrados.length}`);

        log(`${dadosFiltrados.length} linhas restantes no resultado final.`);

        // --- Final: Salvar o arquivo final ---
        if (dadosFiltrados.length === 0) {
            log('AVISO: Nenhum dado correspondeu a todos os critérios de filtro. O arquivo final terá apenas cabeçalhos.');
        }

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
