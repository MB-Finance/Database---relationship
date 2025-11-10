const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = 3000;

// Configuração do Multer para fazer upload dos arquivos em memória (como buffers)
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        // Aceita apenas arquivos excel
        if (file.mimetype.includes('excel') || file.mimetype.includes('spreadsheetml')) {
            cb(null, true);
        } else {
            cb(new Error('Apenas arquivos .xlsx, .xls são permitidos!'), false);
        }
    }
});

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


// ---------- FUNÇÃO: PIPELINE COMPLETO (gera elegiveis_auto.xlsx com colunas preenchidas) ----------
function runFullPipeline(fileBuffers) {
    serverLogs = []; // reset logs para execução
    try {
        log('Iniciando pipeline completo...');

        // --- Ler arquivos a partir dos buffers recebidos via upload ---
        const relatorioBuffer = fileBuffers.relatorio[0].buffer;
        const bitrixBuffer = fileBuffers.bitrix[0].buffer;
        const timeBuffer = fileBuffers.time[0].buffer;
        const contatosBuffer = fileBuffers.contatos[0].buffer;

        // --- Ler relatorio.xlsx ---
        log(`Lendo buffer do arquivo de relatório...`);
        const relWb = xlsx.read(relatorioBuffer);
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
        headerRow[5] = 'faturamento';
        // se quiser manter outros cabeçalhos, não mexa; já renomeamos essas 3

        // --- Montar workbook inicial 'elegiveis_auto' em memória com a planilha principal ---
        const elegiveisWb = xlsx.utils.book_new();
        const elegiveisWs = xlsx.utils.aoa_to_sheet(dataComNovasColunas, { cellDates: true });
        xlsx.utils.book_append_sheet(elegiveisWb, elegiveisWs, NOME_PLANILHA_PRINCIPAL);
        log('Planilha principal criada em memória.');

        // --- Ler Bitrix --- (extraindo CNPJ (coluna H), Fase (B), Responsável (E))
        log(`Lendo buffer do arquivo Bitrix...`);
        const bitrixWb = xlsx.read(bitrixBuffer);
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
        log(`Lendo buffer do arquivo de time...`);
        const timeWb = xlsx.read(timeBuffer);
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

        // --- Ler Arquivo CONTATOSBITRIX (faturamento) ---
        log(`Lendo buffer do arquivo CONTATOSBITRIX...`);
        const contatosWb = xlsx.read(contatosBuffer);
        const contatosWs = contatosWb.Sheets[contatosWb.SheetNames[0]];
        const contatosDataJson = xlsx.utils.sheet_to_json(contatosWs, { defval: '' });

        // Montar mapa para faturamento (CNPJ -> Faturamento)
        const mapFaturamento = {};
        const contatosHeaders = Object.keys(contatosDataJson[0] || {});
        const hCnpjContatos = contatosHeaders.find(h => h && h.toUpperCase().includes('CNPJ')) || contatosHeaders[0];
        const hFaturamento = contatosHeaders[1]; // Coluna B

        for (const row of contatosDataJson) {
            const cnpjKey = normCnpjKey(row[hCnpjContatos]);
            if (cnpjKey) mapFaturamento[cnpjKey] = row[hFaturamento];
        }
        log('Mapa de faturamento criado.');

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
        let countFaturamentoNotFound = 0;
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
            let faturamento = 'Não encontrado';

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

            // lookup do faturamento
            if (cnpjKey && mapFaturamento[cnpjKey] !== undefined) {
                faturamento = mapFaturamento[cnpjKey];
            } else {
                countFaturamentoNotFound++;
            }

            // Retornar novo objeto com as colunas adicionadas (mantendo todas as colunas originais)
            return {
                ...row,
                'fase': fase,
                'responsavel': responsavel,
                'Supervisor': supervisor,
                'faturamento': faturamento
            };
        });

        log(`Lookups concluídos. CNPJs não encontrados (fase/resp): ${countCnpjNotFound}. Responsáveis sem supervisor: ${countRespNotFound}. CNPJs sem faturamento: ${countFaturamentoNotFound}.`);

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

// ---------- ROTAS DO SERVIDOR EXPRESS ----------

// Servir a página HTML principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Rota para executar o pipeline, agora recebendo os arquivos
app.post('/run', upload.fields([
    { name: 'relatorio', maxCount: 1 },
    { name: 'bitrix', maxCount: 1 },
    { name: 'time', maxCount: 1 },
    { name: 'contatos', maxCount: 1 }
]), (req, res) => {

    if (!req.files || !req.files.relatorio || !req.files.bitrix || !req.files.time) {
        return res.status(400).json({
            message: 'Erro: Todos os três arquivos são obrigatórios.',
            logs: ['Requisição recebida, mas um ou mais arquivos não foram enviados.']
        });
    }

    // Os arquivos estão em req.files e seus buffers em req.files.<fieldname>[0].buffer
    const result = runFullPipeline(req.files);

    if (result.success) {
        res.status(200).json({
            message: 'Pipeline executado com sucesso!',
            logs: result.logs
        });
    } else {
        res.status(500).json({
            message: 'Ocorreu um erro durante a execução do pipeline.',
            logs: result.logs,
            error: result.error ? result.error.message : 'Erro desconhecido'
        });
    }
});

// ---------- INICIALIZAÇÃO DO SERVIDOR ----------
app.listen(PORT, () => {
    console.log(`Servidor rodando em http://localhost:${PORT}`);
    console.log('Acesse a página no navegador para fazer o upload dos arquivos e executar o processo.');
});
