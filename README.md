# Sistema de Levantamento Estatístico e Acompanhamento de Documentos

Este projeto é um sistema desenvolvido para auxiliar na análise, filtragem e tratamento de dados de planilhas de registros de chegada de documentos da Consultoria Jurídica (Conjur) da Polícia Militar do Pará (PMPA). O sistema gera relatórios imprimíveis com base nos dados processados.

## Funcionalidades

- **Acesso a Planilhas**: O sistema acessa planilhas do Google Sheets para obter dados de registros de documentos.
- **Filtragem de Dados**: Filtra e trata os dados para identificar prazos, pendências e outras informações relevantes.
- **Geração de Relatórios**: Produz relatórios em formato HTML que podem ser impressos ou visualizados diretamente no navegador.
- **Detecção de Pendências**: Identifica processos que estão pendentes de providências com base em prazos e datas de entrada.
- **Cálculo de Prazos**: Calcula datas de saída máxima com base nos prazos definidos para cada processo.

## Como Usar

1. **Configuração Inicial**:
   - Substitua o valor de `spreadsheetId` nas funções `pegador` e `encontrarUltimaLinhaPreenchida` pelo ID da sua planilha do Google Sheets.
   - Certifique-se de que as abas e colunas da planilha estejam nomeadas corretamente para que o sistema possa acessar os dados.

2. **Execução**:
   - Execute a função `gerar_Relatorio()` para gerar o relatório. Isso abrirá uma janela modal no Google Sheets com o relatório em formato HTML.
   - O relatório pode ser impresso diretamente da janela modal.

3. **Funções Principais**:
   - `pegador(range, aba)`: Acessa uma célula específica na planilha.
   - `encontrarUltimaLinhaPreenchida(aba, coluna)`: Encontra a última linha preenchida em uma coluna específica.
   - `data_chegada(aba, coluna)`: Retorna a última data de chegada de processos em uma aba específica.
   - `pendencia(aba)`: Verifica se há pendências em uma aba específica.
   - `gerar_Relatorio()`: Gera o relatório final com os dados processados.

## Estrutura do Código

- **Funções de Acesso a Dados**:
  - `pegador`, `encontrarUltimaLinhaPreenchida`, `endereco_ultima_linha`, `data_chegada`, `filtrarLinhasNumericas`.

- **Funções de Manipulação de Datas**:
  - `datador`, `getDataDeHoje`, `subtrairDatas`, `somarDias`.

- **Funções de Geração de Relatórios**:
  - `linha_relatorio`, `gerar_Relatorio`, `gerarPaginaHTML`, `gerarTabelaHTML`.

- **Funções Auxiliares**:
  - `mes`, `nada`.

## Exemplo de Uso

```javascript
function exemploUso() {
  // Gera o relatório
  gerar_Relatorio();
}
```

## Requisitos

- Google Apps Script: O código é executado no ambiente do Google Apps Script, que permite a integração com Google Sheets e outras ferramentas do Google.
- Planilha do Google Sheets: O sistema depende de uma planilha do Google Sheets para acessar os dados.

## Autor

- **Gustavo Maciel**
  - GitHub: [@quietbytesilence](https://github.com/quietbytesilence)
  - Telegram: [t/me/lievasomal](https://t.me/lievasomal)

## Licença

Padrão GitHub

