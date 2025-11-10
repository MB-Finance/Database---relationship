/**
 * app.js
 *
 * Script unificado:
 * - Roda como servidor HTTP (porta 3000). POST /run -> executa pipeline completo (ler relatorio.xlsx, bitrix, time, gerar elegiveis_auto.xlsx).
 * - Também oferece modo "procv-only" (se existir elegiveis_auto.xlsx ele gera relatorio_final.xlsx) quando executado diretamente (node app.js).
 *
 * Dependências:
 *   npm install xlsx
 *
 * Coloque os arquivos na mesma pasta:
 *   - relatorio.xlsx
 *   - baixada_do_bitrix.xlsx
 *   - time_novembro.xlsx
 *   (ou apenas elegiveis_auto.xlsx caso queira só o passo PROCV-only)
 *
 * Execução:
 *   node app.js        -> roda em modo CLI: se achar elegiveis_auto.xlsx executa PROCV-only; senão executa pipeline completo (se tiver os arquivos).
 *   Start server: node app.js && abrir http://localhost:3000 e clicar/run POST -> chama pipeline completo.
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// ---------- CONFIGURAÇÃO DE ARQUIVOS ----------
const ARQUIVO_RELATORIO = 'relatorio.xlsx';
const ARQUIVO_ELEGIVEIS = 'elegiveis_auto.xlsx';
const ARQUIVO_BITRIX = 'baixada_do_bitrix.xlsx';
const ARQUIVO_TIME = 'time_novembro.xlsx';

const ARQUIVO_SAIDA_PROCV = 'relatorio_final.xlsx';

const NOME_PLANILHA_PRINCIPAL = 'Sheet1'; // planilha principal que criamos
const NOME_PLANILHA_RELACIONAMENTO = 'C6 - Relacionamento';
const NOME_PLANILHA_SUPERVISORES = 'supervisores';

// ---------- LOGS ----------
let serverLogs = [];
const log = (msg) => {
    const texto = `[${new Date().toISOString()}] ${msg}`;
    console.log(texto);
    serverLogs.push(texto);
};

// ---------- UTILITÁRIOS DE NORMALIZAÇÃO ----------
const norm = v => (v === null || v === undefined) ? '' : String(v).trim();
const normKey = v => norm(v).toUpperCase();
const normCnpjKey = v => norm(v).replace(/\D/g, ''); // mantém só números

// ---------- FUNÇÃO: CHECAR EXISTÊNCIA ARQUIVOS (para pipeline completo) ----------
function checkRequiredFilesForPipeline() {
    const files = [ARQUIVO_RELATORIO, ARQUIVO_BITRIX, ARQUIVO_TIME];
    for (const f of files) {
        if (!fs.existsSync(f)) {
            log(`Arquivo necessário não encontrado: ${f}`);
            return false;
        }
    }
    return true;
}

// ---------- FUNÇÃO: PIPELINE COMPLETO (gera elegiveis_auto.xlsx com colunas preenchidas) ----------
function runFullPipeline() {
    serverLogs = []; // reset logs para execução
    try {
        log('Iniciando pipeline completo...');

        if (!checkRequiredFilesForPipeline()) {
            log('Faltam arquivos para executar o pipeline completo. Abortando pipeline.');
            return { success: false, logs: serverLogs };
        }

        // --- Ler relatorio.xlsx ---
        log(`Lendo ${ARQUIVO_RELATORIO}...`);
        const relWb = xlsx.readFile(ARQUIVO_RELATORIO);
        const relFirstSheet = relWb.Sheets[relWb.SheetNames[0]];
        const relDataAoA = xlsx.utils.sheet_to_json(relFirstSheet, { header: 1, defval: null });

        if (relDataAoA.length === 0) {
            log('Relatório vazio. Abortando.');
            return { success: false, logs: serverLogs };
        }

        // --- Inserir 4 colunas em branco a partir da coluna C (índice 2) ---
        log('Inserindo 4 colunas em branco (C, D, E, F) e renomeando cabeçalhos...');
        const dataComNovasColunas = relDataAoA.map((row, rowIndex) => {
            const newRow = [...row];
            // garante que tenha ao menos 2 posições para inserir no índice 2
            while (newRow.length < 2) newRow.push(null);
            newRow.splice(2, 0, null, null, null, null);
            return newRow;
        });

        // renomeia cabeçalhos das novas colunas imediatamente
        const headerRow = dataComNovasColunas[0];
        headerRow[2] = 'fase';
        headerRow[3] = 'responsavel';
        headerRow[4] = 'Supervisor';
        // se quiser manter outros cabeçalhos, não mexa; já renomeamos essas 3

        // --- Montar workbook inicial 'elegiveis_auto' em memória com a planilha principal ---
        const elegiveisWb = xlsx.utils.book_new();
        const elegiveisWs = xlsx.utils.aoa_to_sheet(dataComNovasColunas, { cellDates: true });
        xlsx.utils.book_append_sheet(elegiveisWb, elegiveisWs, NOME_PLANILHA_PRINCIPAL);
        log('Planilha principal criada em memória.');

        // --- Ler Bitrix --- (extraindo CNPJ (coluna H), Fase (B), Responsável (E))
        log(`Lendo ${ARQUIVO_BITRIX} (Bitrix)...`);
        const bitrixWb = xlsx.readFile(ARQUIVO_BITRIX);
        const bitrixWs = bitrixWb.Sheets[bitrixWb.SheetNames[0]];
        const bitrixDataAoA = xlsx.utils.sheet_to_json(bitrixWs, { header: 1, defval: null });

        // Pegar índices por letras (B=1, E=4, H=7)
        const idxB = xlsx.utils.decode_col('B');
        const idxE = xlsx.utils.decode_col('E');
        const idxH = xlsx.utils.decode_col('H');

        // Extrair dados e montar bitrixColunas com cabeçalho padronizado
        const bitrixRows = [];
        bitrixRows.push(['CNPJ', 'Fase', 'Responsavel']);
        for (let i = 1; i < bitrixDataAoA.length; i++) { // pula cabeçalho original
            const r = bitrixDataAoA[i];
            const cnpjVal = r[idxH];
            const faseVal = r[idxB];
            const respVal = r[idxE];
            // se quiser filtrar linhas sem CNPJ, pode pular
            if (cnpjVal === null || cnpjVal === undefined || String(cnpjVal).trim() === '') continue;
            bitrixRows.push([cnpjVal, faseVal, respVal]);
        }
        const relacionamentoWs = xlsx.utils.aoa_to_sheet(bitrixRows);
        xlsx.utils.book_append_sheet(elegiveisWb, relacionamentoWs, NOME_PLANILHA_RELACIONAMENTO);
        log(`Planilha "${NOME_PLANILHA_RELACIONAMENTO}" adicionada (dados do Bitrix).`);

        // --- Ler Arquivo TIME (supervisores) ---
        log(`Lendo ${ARQUIVO_TIME} (time)...`);
        const timeWb = xlsx.readFile(ARQUIVO_TIME);
        const timeWs = timeWb.Sheets[timeWb.SheetNames[0]];
        // lê como JSON (primeira linha cabeçalho)
        const timeDataJson = xlsx.utils.sheet_to_json(timeWs, { defval: '' });

        // tenta detectar quais colunas são Consultor e Equipe (caso header possua variações)
        const timeHeaders = Object.keys(timeDataJson[0] || {});
        const hConsultor = timeHeaders.find(h => h && h.toUpperCase().includes('CONSULTOR')) || timeHeaders[0];
        const hEquipe = timeHeaders.find(h => h && h.toUpperCase().includes('EQUIPE')) || timeHeaders[1] || timeHeaders[0];

        // monta array com header padronizado
        const supervisoresRows = [['Consultor', 'Equipe']];
        for (const row of timeDataJson) {
            supervisoresRows.push([row[hConsultor] || '', row[hEquipe] || '']);
        }
        const supervisoresWs = xlsx.utils.aoa_to_sheet(supervisoresRows);
        xlsx.utils.book_append_sheet(elegiveisWb, supervisoresWs, NOME_PLANILHA_SUPERVISORES);
        log(`Planilha "${NOME_PLANILHA_SUPERVISORES}" adicionada (dados de time).`);

        // --- Converter a planilha principal (com as 4 colunas em branco) em JSON para filtrar conforme regras ---
        log('Convertendo planilha principal para JSON para aplicar filtros...');
        const elegiveisWsJson = xlsx.utils.sheet_to_json(elegiveisWs, { defval: null });
        if (elegiveisWsJson.length === 0) {
            log('Aviso: a planilha principal gerada está vazia. Abortando.');
            return { success: false, logs: serverLogs };
        }

        // Função auxiliar para procurar cabeçalhos de forma flexível (com fallback por letra)
        const findHeader = (headers, primaryName, fallbackColumnLetter) => {
            let header = headers.find(h => h && h.trim().toUpperCase() === primaryName.toUpperCase());
            if (header) return header;
            // fallback por letra (se existir)
            const colIndex = xlsx.utils.decode_col(fallbackColumnLetter);
            if (headers[colIndex]) {
                log(`AVISO: Coluna "${primaryName}" não encontrada pelo nome. Usando fallback '${fallbackColumnLetter}' (${headers[colIndex]}).`);
                return headers[colIndex];
            }
            return null;
        };

        // pegar headers e determinar colunas para filtros (ajuste conforme seu relatorio original)
        const headersRel = Object.keys(elegiveisWsJson[0]);
        const colElegivel = findHeader(headersRel, 'FL_ELEGIVEL_VENDA_C6PAY', 'AK');
        const colTipoPessoa = findHeader(headersRel, 'TIPO_PESSOA', 'H');
        const colDataAprovacao = findHeader(headersRel, 'DT_APROVACAO_PAY', 'AK');
        const colStatusCC = findHeader(headersRel, 'STATUS_CC', 'Y');

        if (!colElegivel || !colTipoPessoa || !colDataAprovacao || !colStatusCC) {
            log('Erro: não foi possível localizar todas as colunas de filtro (com fallback). Abortando pipeline.');
            return { success: false, logs: serverLogs };
        }

        log(`Total de linhas antes do filtro: ${elegiveisWsJson.length}`);

        // Aplicar filtros:
        const filtro1 = elegiveisWsJson.filter(row => row[colElegivel] === 1 || row[colElegivel] === '1');
        log(`Após filtro FL_ELEGIVEL_VENDA_C6PAY = "1": ${filtro1.length}`);

        const filtro2 = filtro1.filter(row => String(row[colTipoPessoa]).toUpperCase() === 'PJ');
        log(`Após filtro TIPO_PESSOA = "PJ": ${filtro2.length}`);

        const filtro3 = filtro2.filter(row => row[colDataAprovacao] === null || row[colDataAprovacao] === '' || typeof row[colDataAprovacao] === 'undefined');
        log(`Após filtro DT_APROVACAO_PAY vazia: ${filtro3.length}`);

        const dadosFiltrados = filtro3.filter(row => String(row[colStatusCC]).toUpperCase() === 'LIBERADA');
        log(`Após filtro STATUS_CC = "LIBERADA": ${dadosFiltrados.length}`);

        if (dadosFiltrados.length === 0) {
            log('Nenhuma linha restou após os filtros. Gerando arquivo com cabeçalhos apenas.');
            // Gerar arquivo com apenas cabeçalho caso necessário
            const onlyHeaders = [Object.keys(elegiveisWsJson[0])];
            const wbOnly = xlsx.utils.book_new();
            xlsx.utils.book_append_sheet(wbOnly, xlsx.utils.aoa_to_sheet(onlyHeaders), NOME_PLANILHA_PRINCIPAL);
            xlsx.utils.book_append_sheet(wbOnly, relacionamentoWs, NOME_PLANILHA_RELACIONAMENTO);
            xlsx.utils.book_append_sheet(wbOnly, supervisoresWs, NOME_PLANILHA_SUPERVISORES);
            xlsx.writeFile(wbOnly, ARQUIVO_ELEGIVEIS);
            log(`Arquivo ${ARQUIVO_ELEGIVEIS} salvo (somente cabeçalhos).`);
            return { success: true, logs: serverLogs };
        }

        // --- Montar mapas para Lookup (substitui PROCV com fórmulas) ---
        log('Montando mapas de lookup (substituindo PROCV por lógica JS)...');

        // Map do relacionamento (CNPJ -> { fase, responsavel })
        const mapBitrix = {};
        // bitrixRows já foi montado acima com header; slice(1)
        for (let i = 1; i < bitrixRows.length; i++) {
            const r = bitrixRows[i];
            const rawCnpj = r[0];
            if (!rawCnpj) continue;
            const key = normCnpjKey(rawCnpj);
            mapBitrix[key] = { fase: norm(r[1]), responsavel: norm(r[2]) };
        }

        // Map de supervisores (Consultor -> Equipe)
        const mapTime = {};
        for (let i = 1; i < supervisoresRows.length; i++) {
            const [consultorRaw, equipeRaw] = supervisoresRows[i];
            const consultorKey = normKey(consultorRaw);
            if (!consultorKey) continue;
            if (!mapTime[consultorKey]) mapTime[consultorKey] = norm(equipeRaw);
        }

        // --- Preencher colunas 'fase', 'responsavel', 'Supervisor' diretamente nos dados filtrados ---
        log('Executando lookups e preenchendo colunas nas linhas filtradas...');
        let countCnpjNotFound = 0;
        let countRespNotFound = 0;
        const dadosComLookups = dadosFiltrados.map((row) => {
            // A suposição: coluna CNPJ em seu relatorio original pode variar; vamos procurar
            // possíveis chave 'CNPJ' (maiúsc/minúsc) no objeto row
            const possibleCnpjKeys = Object.keys(row).filter(k => k && k.toUpperCase().includes('CNPJ'));
            let rawCnpjValue = '';
            if (possibleCnpjKeys.length > 0) {
                rawCnpjValue = row[possibleCnpjKeys[0]];
            } else {
                // tentar por posição: se existir chave parecida
                rawCnpjValue = row['CNPJ'] || row['cnpj'] || '';
            }
            const cnpjKey = normCnpjKey(rawCnpjValue);

            let fase = 'Não encontrado';
            let responsavel = 'Não encontrado';
            let supervisor = 'Não encontrado';

            if (cnpjKey && mapBitrix[cnpjKey]) {
                fase = mapBitrix[cnpjKey].fase || 'Não encontrado';
                responsavel = mapBitrix[cnpjKey].responsavel || 'Não encontrado';
            } else {
                countCnpjNotFound++;
            }

            // lookup do supervisor a partir do responsavel (consultor)
            const respKey = normKey(responsavel);
            if (respKey && mapTime[respKey]) {
                supervisor = mapTime[respKey];
            } else {
                if (responsavel !== 'Não encontrado') countRespNotFound++;
            }

            // Retornar novo objeto com as colunas adicionadas (mantendo todas as colunas originais)
            return {
                ...row,
                'fase': fase,
                'responsavel': responsavel,
                'Supervisor': supervisor
            };
        });

        log(`Lookups concluídos. CNPJs não encontrados: ${countCnpjNotFound}. Responsáveis (consultor) sem supervisor: ${countRespNotFound}.`);

        // --- Converter dadosComLookups para sheet (inclui cabeçalho) e salvar no arquivo elegiveis_auto.xlsx ---
        log(`Salvando arquivo ${ARQUIVO_ELEGIVEIS} com dados preenchidos...`);
        const finalSheet = xlsx.utils.json_to_sheet(dadosComLookups, { skipHeader: false });
        const finalWb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(finalWb, finalSheet, NOME_PLANILHA_PRINCIPAL);
        xlsx.utils.book_append_sheet(finalWb, relacionamentoWs, NOME_PLANILHA_RELACIONAMENTO);
        xlsx.utils.book_append_sheet(finalWb, supervisoresWs, NOME_PLANILHA_SUPERVISORES);
        xlsx.writeFile(finalWb, ARQUIVO_ELEGIVEIS);

        log(`Arquivo ${ARQUIVO_ELEGIVEIS} salvo com sucesso. Pipeline completo finalizado.`);
        return { success: true, logs: serverLogs };

    } catch (err) {
        log(`Erro no pipeline: ${err.message}`);
        log(err.stack);
        return { success: false, logs: serverLogs, error: err };
    }
}

// ---------- FUNÇÃO: PROCV-ONLY (le elegiveis_auto.xlsx e gera relatorio_final.xlsx) ----------
function runProcvOnly() {
    serverLogs = [];
    try {
        log('Iniciando PROCV-only (ler elegiveis_auto.xlsx -> gerar relatorio_final.xlsx)...');

        if (!fs.existsSync(ARQUIVO_ELEGIVEIS)) {
            log(`Arquivo ${ARQUIVO_ELEGIVEIS} não encontrado. Abortando PROCV-only.`);
            return { success: false, logs: serverLogs };
        }

        const wb = xlsx.readFile(ARQUIVO_ELEGIVEIS);
        const wsPrincipal = wb.Sheets[NOME_PLANILHA_PRINCIPAL] || wb.Sheets[wb.SheetNames[0]];
        const wsRelacionamento = wb.Sheets[NOME_PLANILHA_RELACIONAMENTO] || wb.Sheets[Object.keys(wb.Sheets).find(n => n.toLowerCase().includes('relacion'))];
        const wsSupervisor = wb.Sheets[NOME_PLANILHA_SUPERVISORES] || wb.Sheets[Object.keys(wb.Sheets).find(n => n.toLowerCase().includes('supervis'))];

        if (!wsPrincipal) {
            log(`Planilha principal "${NOME_PLANILHA_PRINCIPAL}" não encontrada em ${ARQUIVO_ELEGIVEIS}. Abortando.`);
            return { success: false, logs: serverLogs };
        }
        if (!wsRelacionamento) {
            log(`Planilha relacionamento não encontrada em ${ARQUIVO_ELEGIVEIS}. Abortando.`);
            return { success: false, logs: serverLogs };
        }
        if (!wsSupervisor) {
            log(`Planilha supervisor não encontrada em ${ARQUIVO_ELEGIVEIS}. Abortando.`);
            return { success: false, logs: serverLogs };
        }

        // Ler como JSON (usa cabeçalhos)
        const dadosPrincipal = xlsx.utils.sheet_to_json(wsPrincipal, { defval: '' });
        const dadosRelacionamento = xlsx.utils.sheet_to_json(wsRelacionamento, { defval: '' });
        const dadosSupervisor = xlsx.utils.sheet_to_json(wsSupervisor, { defval: '' });

        // Montar mapas (com normalização)
        const mapaRelacionamento = new Map();
        for (const linha of dadosRelacionamento) {
            // tenta detectar o nome da coluna CNPJ independentemente do case
            const keys = Object.keys(linha);
            const keyCnpj = keys.find(k => k && k.toUpperCase().includes('CNPJ')) || keys[0];
            const keyFase = keys.find(k => k && k.toUpperCase().includes('FASE')) || keys[1];
            const keyResponsavel = keys.find(k => k && k.toUpperCase().includes('RESPONS')) || keys[2];

            const cnpj = linha[keyCnpj];
            const fase = linha[keyFase];
            const responsavel = linha[keyResponsavel];

            if (cnpj) {
                mapaRelacionamento.set(normCnpjKey(cnpj), { fase: norm(fase), responsavel: norm(responsavel) });
            }
        }

        const mapaSupervisor = new Map();
        for (const linha of dadosSupervisor) {
            const keys = Object.keys(linha);
            const keyCons = keys.find(k => k && k.toUpperCase().includes('CONSULT')) || keys[0];
            const keyEquipe = keys.find(k => k && k.toUpperCase().includes('EQUIPE')) || keys[1];
            const consultor = linha[keyCons];
            const equipe = linha[keyEquipe];
            if (consultor) mapaSupervisor.set(normKey(consultor), norm(equipe));
        }

        log('Mapas do PROCV-only criados.');

        // Processar cada linha principal e adicionar campos
        const resultadoFinal = [];
        let notFoundCnpj = 0;
        let notFoundSupervisor = 0;

        for (const linha of dadosPrincipal) {
            // detectar campo cnpj na linha
            const keys = Object.keys(linha);
            const keyCnpj = keys.find(k => k && k.toUpperCase().includes('CNPJ')) || keys[0];
            const rawCnpj = linha[keyCnpj];

            let fase = 'Não encontrado';
            let responsavel = 'Não encontrado';
            let supervisor = 'Não encontrado';

            if (rawCnpj) {
                const cnpjKey = normCnpjKey(rawCnpj);
                if (mapaRelacionamento.has(cnpjKey)) {
                    const dadosRel = mapaRelacionamento.get(cnpjKey);
                    fase = dadosRel.fase || 'Não encontrado';
                    responsavel = dadosRel.responsavel || 'Não encontrado';
                } else {
                    notFoundCnpj++;
                }
            } else {
                notFoundCnpj++;
            }

            if (responsavel !== 'Não encontrado' && mapaSupervisor.has(normKey(responsavel))) {
                supervisor = mapaSupervisor.get(normKey(responsavel)) || 'Não encontrado';
            } else {
                if (responsavel !== 'Não encontrado') notFoundSupervisor++;
            }

            // adiciona sem sobrescrever campos originais
            resultadoFinal.push({
                ...linha,
                'fase': fase,
                'responsavel': responsavel,
                'supervisor': supervisor
            });
        }

        log(`PROCV-only: CNPJs não encontrados: ${notFoundCnpj}. Responsáveis sem supervisor: ${notFoundSupervisor}.`);

        // Criar nova planilha/arquivo de saída
        const novaWs = xlsx.utils.json_to_sheet(resultadoFinal);
        const novoWb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(novoWb, novaWs, 'Resultado');
        xlsx.writeFile(novoWb, ARQUIVO_SAIDA_PROCV);

        log(`Arquivo ${ARQUIVO_SAIDA_PROCV} criado com sucesso.`);
        return { success: true, logs: serverLogs };

    } catch (err) {
        log(`Erro no PROCV-only: ${err.message}`);
        log(err.stack);
        return { success: false, logs: serverLogs, error: err };
    }
}

// ---------- HTTP SERVER (rota simples + logs) ----------
const server = http.createServer((req, res) => {
    if (req.method === 'POST' && req.url === '/run') {
        // Executa pipeline completo
        const result = runFullPipeline();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ message: 'Pipeline executado', result }));
        return;
    }

    if (req.method === 'GET' && req.url === '/logs') {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ logs: serverLogs }));
        return;
    }

    // Página simples para acionar o /run via form
    if (req.method === 'GET' && req.url === '/') {
        const html = `
            <html>
                <head><meta charset="utf-8"><title>Automação XLSX</title></head>
                <body>
                    <h2>Automação XLSX</h2>
                    <p>POST <code>/run</code> para executar o pipeline completo.</p>
                    <form method="post" action="/run">
                        <button type="submit">Executar pipeline completo</button>
                    </form>
                    <p>GET <code>/logs</code> para ver logs da última execução.</p>
                </body>
            </html>
        `;
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(html);
        return;
    }

    res.writeHead(404);
    res.end('Not Found');
});

// Se executado diretamente (CLI), decide que fluxo executar com base nos arquivos presentes.
if (require.main === module) {
    // Se tiver parametro de CLI "procv" -> roda PROCV-only
    const cliArg = process.argv[2];
    if (cliArg && (cliArg.toLowerCase() === 'procv' || cliArg.toLowerCase() === 'procv-only')) {
        const r = runProcvOnly();
        console.log('PROCV-only finalizado. Logs:');
        console.log(r.logs.join('\n'));
        process.exit(r.success ? 0 : 1);
    }

    // Se existir elegiveis_auto.xlsx -> por padrão executa PROCV-only
    if (fs.existsSync(ARQUIVO_ELEGIVEIS) && !checkRequiredFilesForPipeline()) {
        const r = runProcvOnly();
        console.log('Executado PROCV-only (detecção automática). Logs:');
        console.log(r.logs.join('\n'));
        process.exit(r.success ? 0 : 1);
    }

    // Caso contrário, tenta executar o pipeline completo localmente (se arquivos estiverem presentes)
    if (checkRequiredFilesForPipeline()) {
        const r = runFullPipeline();
        console.log('Pipeline completo finalizado. Logs:');
        console.log(r.logs.join('\n'));
        process.exit(r.success ? 0 : 1);
    }

    // Se nenhum dos cenários aplicou, inicializa o servidor HTTP e aguarda POST /run
    const PORT = 3000;
    server.listen(PORT, () => {
        console.log(`Servidor rodando em http://localhost:${PORT}`);
        console.log('Coloque os arquivos .xlsx na mesma pasta e acesse o servidor para executar o pipeline.');
    });
}
