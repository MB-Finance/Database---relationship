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

let serverLogs = []; // Array para armazenar logs do servidor

// Função para registrar logs
const log = (message) => {
    console.log(message);
    serverLogs.push(message); // Armazena o log para enviar ao front-end
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

        // **CORREÇÃO CRÍTICA**: Renomeia os cabeçalhos das novas colunas IMEDIATAMENTE.
        const headerRow = dataComNovasColunas[0];
        headerRow[2] = 'fase';
        headerRow[3] = 'responsável';
        headerRow[4] = 'Supervisor';

        const elegiveisAutoWb = xlsx.utils.book_new();
        const elegiveisAutoWs = xlsx.utils.aoa_to_sheet(dataComNovasColunas, { cellDates: true });
        xlsx.utils.book_append_sheet(elegiveisAutoWb, elegiveisAutoWs, NOME_PLANILHA_PRINCIPAL);
        log('Planilha principal de "elegiveis_auto" criada na memória.');

        // --- Passo 3, 4, 5, 6: Processar planilhas de apoio ---
        log(`Lendo ${ARQUIVO_BITRIX}...`);
        const bitrixWb = xlsx.readFile(ARQUIVO_BITRIX);
        const bitrixWs = bitrixWb.Sheets[bitrixWb.SheetNames[0]];
        // Lê os dados como uma matriz para usar os índices das colunas diretamente
        const bitrixData = xlsx.utils.sheet_to_json(bitrixWs, { header: 1 }); 

        // Pega os índices das colunas especificadas (B=1, E=4, H=7)
        const colIndexCnpj = xlsx.utils.decode_col('H');
        const colIndexFase = xlsx.utils.decode_col('B');
        const colIndexResponsavel = xlsx.utils.decode_col('E');

        // Extrai as colunas CNPJ (H), Fase (B) e Responsável (E), incluindo o cabeçalho
        const bitrixColunasData = bitrixData.slice(1).map((row) => [ // .slice(1) para pular o cabeçalho original
            row[colIndexCnpj], 
            row[colIndexFase], 
            row[colIndexResponsavel]
        ]);
        const bitrixColunas = [['CNPJ', 'Fase', 'Responsavel'], ...bitrixColunasData]; // Adiciona o cabeçalho padronizado no início

        const relacionamentoWs = xlsx.utils.aoa_to_sheet(bitrixColunas);
        xlsx.utils.book_append_sheet(elegiveisAutoWb, relacionamentoWs, NOME_PLANILHA_RELACIONAMENTO);
        log(`Planilha "${NOME_PLANILHA_RELACIONAMENTO}" adicionada.`);

        log(`Lendo ${ARQUIVO_TIME}...`);
        const timeWb = xlsx.readFile(ARQUIVO_TIME);
        const timeWs = timeWb.Sheets[timeWb.SheetNames[0]];
        const timeData = xlsx.utils.sheet_to_json(timeWs); // Lê usando a primeira linha como cabeçalho
        
        const timeHeaders = Object.keys(timeData[0] || {});
        const hConsultor = timeHeaders.find(h => h.toUpperCase().includes('CONSULTOR'));
        const hEquipe = timeHeaders.find(h => h.toUpperCase().includes('EQUIPE'));

        // Extrai os dados e adiciona a linha de cabeçalho manualmente
        let timeColunas = timeData.map(row => [row[hConsultor] || '', row[hEquipe] || '']); // Garante que valores nulos sejam strings vazias
        timeColunas.unshift(['Consultor', 'Equipe']); // Adiciona o cabeçalho no início do array

        const supervisoresWs = xlsx.utils.aoa_to_sheet(timeColunas);
        xlsx.utils.book_append_sheet(elegiveisAutoWb, supervisoresWs, NOME_PLANILHA_SUPERVISORES);
        log(`Planilha "${NOME_PLANILHA_SUPERVISORES}" adicionada.`);

        // --- Passo 8: Filtros e exclusão de linhas ---
        log('Aplicando filtros e removendo linhas...');
        let finalData = xlsx.utils.sheet_to_json(elegiveisAutoWs, { defval: null });
        
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
        const colDataAprovacao = findHeader(headers, 'DT_APROVACAO_PAY', 'AK'); // Usando AK como fallback, conforme prompt anterior
        const colStatusCC = findHeader(headers, 'STATUS_CC', 'Y'); //
        
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

        const dadosFiltrados = filtro3.filter(row => String(row[colStatusCC]).toUpperCase() === 'LIBERADA'); //
        log(`Linhas restantes após filtro STATUS_CC = "LIBERADA": ${dadosFiltrados.length}`);

        log(`${dadosFiltrados.length} linhas restantes no resultado final.`);

        // --- Passo 7: Lógica de PROCV em etapas separadas ---
        log('Iniciando buscas PROCV em etapas para garantir a ordem de cálculo...');

        // Prepara os dados filtrados para receber as fórmulas.
        // Converte o JSON de volta para um array de arrays (aoa) para manipulação de células.
        let dadosParaProcv = [Object.keys(dadosFiltrados[0] || {}), ...dadosFiltrados.map(Object.values)];

        // Cria a pasta de trabalho que será usada para os cálculos.
        const wbParaCalculo = xlsx.utils.book_new();
        // Adiciona as planilhas de dependência (lookups) desde o início.
        xlsx.utils.book_append_sheet(wbParaCalculo, relacionamentoWs, NOME_PLANILHA_RELACIONAMENTO);
        xlsx.utils.book_append_sheet(wbParaCalculo, supervisoresWs, NOME_PLANILHA_SUPERVISORES);

        // ETAPA 1: Aplicar PROCV para 'fase' e 'responsável'.
        log("Etapa 1: Aplicando fórmulas para 'fase' e 'responsável'...");
        for (let i = 1; i < dadosParaProcv.length; i++) { // Começa em 1 para pular o cabeçalho
            const rowIndex = i + 1; // O índice da linha no Excel é baseado em 1 e inclui o cabeçalho.
            dadosParaProcv[i][2] = { f: `IFERROR(VLOOKUP(B${rowIndex},'${NOME_PLANILHA_RELACIONAMENTO}'!A:B,2,FALSE),"Não encontrado")` };
            dadosParaProcv[i][3] = { f: `IFERROR(VLOOKUP(B${rowIndex},'${NOME_PLANILHA_RELACIONAMENTO}'!A:C,3,FALSE),"Não encontrado")` };
        }
        
        // Adiciona a planilha principal com as fórmulas da Etapa 1.
        const wsComFormulas1 = xlsx.utils.aoa_to_sheet(dadosParaProcv);
        xlsx.utils.book_append_sheet(wbParaCalculo, wsComFormulas1, NOME_PLANILHA_PRINCIPAL);
        
        // Força o cálculo das fórmulas escrevendo e lendo o buffer na memória.
        const buffer1 = xlsx.write(wbParaCalculo, { type: 'buffer', bookType: 'xlsx' });
        const wbCalculado1 = xlsx.read(buffer1, { type: 'buffer', cellFormula: false });
        const wsCalculado1 = wbCalculado1.Sheets[NOME_PLANILHA_PRINCIPAL];
        const dadosCalculados1 = xlsx.utils.sheet_to_json(wsCalculado1, { header: 1 });
        log("Etapa 1: Valores de 'fase' e 'responsável' calculados.");

        // ETAPA 2: Aplicar PROCV para 'Supervisor' usando os dados calculados da Etapa 1.
        log("Etapa 2: Aplicando fórmulas para 'Supervisor'...");
        for (let i = 1; i < dadosCalculados1.length; i++) { // Começa em 1 para pular o cabeçalho
            const rowIndex = i + 1;
            dadosCalculados1[i][4] = { f: `IFERROR(VLOOKUP(D${rowIndex},'${NOME_PLANILHA_SUPERVISORES}'!A:B,2,FALSE),"Não encontrado")` };
        }

        // Atualiza a planilha principal na pasta de trabalho já existente com as novas fórmulas.
        wbParaCalculo.Sheets[NOME_PLANILHA_PRINCIPAL] = xlsx.utils.aoa_to_sheet(dadosCalculados1);
        
        // Força o cálculo final usando a mesma técnica de buffer.
        const buffer2 = xlsx.write(wbParaCalculo, { type: 'buffer', bookType: 'xlsx' });
        const wbCalculadoFinal = xlsx.read(buffer2, { type: 'buffer', cellFormula: false });
        const finalCalculatedData = xlsx.utils.sheet_to_json(wbCalculadoFinal.Sheets[NOME_PLANILHA_PRINCIPAL], { header: 1 });
        log("Etapa 2: Valores de 'Supervisor' calculados.");

        // Os dados em 'finalCalculatedData' já estão no formato de array de arrays, incluindo o cabeçalho.
        const finalDataRows = finalCalculatedData.slice(1); // Pega apenas as linhas de dados

        // --- Final: Salvar o arquivo final ---
        const finalWb = xlsx.utils.book_new();
        // Reutiliza as planilhas já calculadas do workbook final.
        xlsx.utils.book_append_sheet(finalWb, wbCalculadoFinal.Sheets[NOME_PLANILHA_PRINCIPAL], NOME_PLANILHA_PRINCIPAL);
        xlsx.utils.book_append_sheet(finalWb, wbCalculadoFinal.Sheets[NOME_PLANILHA_RELACIONAMENTO], NOME_PLANILHA_RELACIONAMENTO);
        xlsx.utils.book_append_sheet(finalWb, wbCalculadoFinal.Sheets[NOME_PLANILHA_SUPERVISORES], NOME_PLANILHA_SUPERVISORES);

        if (finalDataRows.length === 0 && dadosFiltrados.length > 0) {
            log('AVISO: Nenhum dado restou após todos os processos. O arquivo final pode estar vazio ou conter apenas cabeçalhos.');
        }

        xlsx.writeFile(finalWb, ARQUIVO_ELEGIVEIS);
        log(`Processo concluído! Arquivo "${ARQUIVO_ELEGIVEIS}" salvo com sucesso.`);

    } catch (error) {
        log(`Ocorreu um erro: ${error.message}`);
        console.error(error);
    }
}

// Servidor HTTP simples para servir a página HTML
const server = http.createServer((req, res) => {
    if (req.method === 'POST' && req.url === '/run') {
        serverLogs = []; // Limpa os logs antigos a cada nova execução
        runAutomation(); //
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ message: 'Processo concluído.', logs: serverLogs })); // Envia todos os logs para o cliente
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
