const xlsx = require('xlsx');

// --- CONFIGURAÇÕES ---
const ARQUIVO_ENTRADA = 'elegiveis_auto.xlsx';
const ARQUIVO_SAIDA = 'relatorio_final.xlsx';

const NOME_PLANILHA_PRINCIPAL = 'Sheet1';
const NOME_PLANILHA_RELACIONAMENTO = 'C6 - relacionamento';
const NOME_PLANILHA_SUPERVISOR = 'supervisor';

// --- FUNÇÃO PRINCIPAL ---
function executarProcv() {
    try {
        console.log(`Iniciando o processo. Lendo o arquivo: ${ARQUIVO_ENTRADA}`);

        // 1. Ler o arquivo de entrada
        const workbook = xlsx.readFile(ARQUIVO_ENTRADA);

        // 2. Validar e carregar as planilhas necessárias
        const wsPrincipal = workbook.Sheets[NOME_PLANILHA_PRINCIPAL];
        const wsRelacionamento = workbook.Sheets[NOME_PLANILHA_RELACIONAMENTO];
        const wsSupervisor = workbook.Sheets[NOME_PLANILHA_SUPERVISOR];

        if (!wsPrincipal) {
            throw new Error(`A planilha "${NOME_PLANILHA_PRINCIPAL}" não foi encontrada.`);
        }
        if (!wsRelacionamento) {
            throw new Error(`A planilha "${NOME_PLANILHA_RELACIONAMENTO}" não foi encontrada.`);
        }
        if (!wsSupervisor) {
            throw new Error(`A planilha "${NOME_PLANILHA_SUPERVISOR}" não foi encontrada.`);
        }

        console.log('Planilhas carregadas com sucesso.');

        // 3. Converter planilhas para JSON para facilitar a manipulação
        const dadosPrincipal = xlsx.utils.sheet_to_json(wsPrincipal);
        const dadosRelacionamento = xlsx.utils.sheet_to_json(wsRelacionamento);
        const dadosSupervisor = xlsx.utils.sheet_to_json(wsSupervisor);

        // 4. Criar mapas de busca para otimizar o PROCV (VLOOKUP)
        // Mapa para relacionamento: { CNPJ -> { fase, responsavel } }
        const mapaRelacionamento = new Map();
        for (const linha of dadosRelacionamento) {
            const cnpj = linha['CNPJ'];
            const fase = linha['Fase'];
            const responsavel = linha['Responsavel'];
            if (cnpj) {
                mapaRelacionamento.set(String(cnpj), { fase, responsavel });
            }
        }

        // Mapa para supervisor: { Responsavel -> Equipe }
        const mapaSupervisor = new Map();
        for (const linha of dadosSupervisor) {
            const consultor = linha['Consultor'];
            const equipe = linha['Equipe'];
            if (consultor) {
                mapaSupervisor.set(String(consultor), equipe);
            }
        }

        console.log('Mapas de busca criados.');

        // 5. Processar os dados da planilha principal
        const resultadoFinal = [];
        for (const linha of dadosPrincipal) {
            const cnpj = linha['CNPJ']; // Assumindo que a coluna de CNPJ se chama 'CNPJ'

            let fase = 'Não encontrado';
            let responsavel = 'Não encontrado';
            let supervisor = 'Não encontrado';

            // PROCV 1 e 2: Buscar fase e responsável
            if (cnpj && mapaRelacionamento.has(String(cnpj))) {
                const dadosRel = mapaRelacionamento.get(String(cnpj));
                fase = dadosRel.fase || 'Não encontrado';
                responsavel = dadosRel.responsavel || 'Não encontrado';
            }

            // PROCV 3: Buscar supervisor usando o responsável encontrado
            if (responsavel !== 'Não encontrado' && mapaSupervisor.has(responsavel)) {
                supervisor = mapaSupervisor.get(responsavel) || 'Não encontrado';
            }

            // Adicionar as novas colunas à linha
            const novaLinha = {
                ...linha,
                'fase': fase,
                'responsável': responsavel,
                'supervisor': supervisor
            };
            resultadoFinal.push(novaLinha);
        }

        console.log('Processamento de dados concluído.');

        // 6. Criar uma nova planilha com os resultados
        const novaPlanilha = xlsx.utils.json_to_sheet(resultadoFinal);

        // 7. Criar um novo workbook e adicionar a nova planilha
        const novoWorkbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(novoWorkbook, novaPlanilha, 'Resultado');

        // 8. Salvar o novo arquivo
        xlsx.writeFile(novoWorkbook, ARQUIVO_SAIDA);

        console.log(`Processo finalizado com sucesso! O arquivo "${ARQUIVO_SAIDA}" foi criado.`);

    } catch (error) {
        console.error('Ocorreu um erro durante o processo:', error.message);
    }
}

// Executar a função
executarProcv();
