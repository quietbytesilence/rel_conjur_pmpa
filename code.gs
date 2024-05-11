/*SOFTWARE DE LEVTAMENTO ESTÍSTIVO E ACOMPANHAMENO DE DOCUMENTOS.
CRIADO POR GUSTAVO MACIEL
GIT: @quietbytesilence
TELEGRAM: t/me/lievasomal
=========================
O sistema analisa a planilha de registros de chegada de documentos da Conjur/PMPA
Analisa os dados, filtra e trata para produzir um relatório imprimível.
*/
  

function pegador(range, aba) {
  var spreadsheetId = "adicionar_seu_id";
  var sheetName = aba;
  var cellValue = 'Indefinido'; // Valor padrão em caso de erro

  try {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);
    var cellValue = sheet.getRange(range).getValue();
  } catch (error) {
    console.error('Erro ao acessar a célula:', error);
  }

  return cellValue;
}

function encontrarUltimaLinhaPreenchida(aba, coluna) {
    var spreadsheetId = "adicionar_seu_id";
    var sheetName = aba;
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);
    var lastRow = sheet.getLastRow();
    var range = coluna + "1:" + coluna + lastRow;
    var values = sheet.getRange(range).getValues();
    var ultimaLinha = 0;

    for (let linha = values.length - 1; linha >= 0; linha--) {
        if (values[linha].some(valor => valor !== '')) {
            ultimaLinha = linha + 1; // Adiciona 1 porque os índices começam em 0
            break;
        }
    }
    return ultimaLinha;
    
}
function endereco_ultima_linha(coluna, aba) {
  var endereco = ((coluna+encontrarUltimaLinhaPreenchida(aba, coluna)))
  return endereco
}


function datador(data) {
  const stringOriginal = data;

  const partes = data.split("/");

  const dia = partes[0];
  const mes = partes[1];
  const ano = partes[2];
  console.log("Dia:", dia); // Saída: Dia: 09
  console.log("Mês:", mes); // Saída: Mês: 05
  console.log("Ano:", ano); // Saída: Ano: 2024

}
function data_chegada(aba, coluna){
  celula = coluna+encontrarUltimaLinhaPreenchida(aba, coluna)
  data = pegador(celula, aba)
  console.log('A última data de chegada de procesos do ' + aba + ' Foi em: ')
  datador(data)
  return data
}

console.log('[---------------]')
console.log('SERVIÇOS INICIADO')
console.log('TESTANDO DATA', getDataDeHoje())
console.log('Ultima Célula Jur I', encontrarUltimaLinhaPreenchida('JUR I','a'))
console.log('Ultima Célula Jur II', encontrarUltimaLinhaPreenchida('JUR II','a'))
console.log('Ultima Célula Jur III', encontrarUltimaLinhaPreenchida('JUR III','a'))
console.log('Ultima Célula Jur IV', encontrarUltimaLinhaPreenchida('JUR IV','a'))
console.log('Funções Criticas finalizadas com sucesso')
//console.log(gerarPaginaHTML([['Nome', 'Idade'],['Gustavo','20'],['Andrey','19']]))

function getDataDeHoje(tipo) {
  const hoje = new Date();
  const dia = hoje.getDate().toString().padStart(2, '0');
  const mes = (hoje.getMonth() + 1).toString().padStart(2, '0'); // Mês começa do zero
  const ano = hoje.getFullYear().toString();
  //return `${dia}/${mes}/${ano}`;
  if (tipo == 'dia'){
    return dia;
  }
  else if (tipo == 'mes'){
    return mes;
  }
  else if (tipo == 'ano'){
    return ano
  }
  else {
    return `${dia}/${mes}/${ano}`
  }
}


function subtrairDatas(data1, data2) {
    var partesData1 = data1.split('/');
    var partesData2 = data2.split('/');
    var dataInicio = new Date(partesData1[2], partesData1[1] - 1, partesData1[0]); // Ano, Mês (zero-based), Dia
    var dataFim = new Date(partesData2[2], partesData2[1] - 1, partesData2[0]); // Ano, Mês (zero-based), Dia
    var diferenca = dataFim - dataInicio;
    var dias = Math.floor(diferenca / (1000 * 60 * 60 * 24));
    return dias;
}
// TESTES DE NOVAS IMPLEMENTAÇÕES

function filtrarLinhasNumericas(aba, coluna) {
  var spreadsheetId = "adicionar_seu_id";
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(aba);
  var range = sheet.getRange(coluna + '1:' + coluna); // Obter o intervalo da coluna
  var values = range.getValues();
  var linhasNumericas = [];

  // Iterar sobre os valores da coluna
  for (var i = 0; i < values.length; i++) {
    var valor = values[i][0];
    // Verificar se o valor é numérico
    if (!isNaN(parseFloat(valor)) && isFinite(valor)) {
      linhasNumericas.push(i + 1); // Adicionar o número da linha se o valor for numérico
    }
  }

  return linhasNumericas;
}
console.log(filtrarLinhasNumericas('JUR III', 'd'))
function somarDias(data, dias) {
  var partesData = data.split('/');
  var dia = parseInt(partesData[0], 10);
  var mes = parseInt(partesData[1], 10);
  var ano = parseInt(partesData[2], 10);
  var novaData = new Date(ano, mes - 1, dia); // Note que o mês é baseado em zero, então subtraímos 1
  novaData.setDate(novaData.getDate() + dias);
  var novoDia = novaData.getDate();
  var novoMes = novaData.getMonth() + 1; // Adicionamos 1 porque o mês é baseado em zero
  var novoAno = novaData.getFullYear();
  var novaDataFormatada = (novoDia < 10 ? '0' : '') + novoDia + '/' + (novoMes < 10 ? '0' : '') + novoMes + '/' + novoAno;
  return novaDataFormatada;
}



function linha_relatorio(aba) {
  prazos = filtrarLinhasNumericas(aba, 'D');
  resposta = [];
  for (var i = 0; i < prazos.length; i++) {
    if (pegador(`P${prazos[i]}`, aba) != 'PROVIDENCIADO') {
      var entrada = pegador(`A${prazos[i]}`, aba);
      var prazo = parseInt(pegador(`d${prazos[i]}`, aba), 10);
      var dataSaidaMaxima = somarDias(entrada, prazo);
      var linhaAtual = [
        pegador(`c${prazos[i]}`, aba), // PAE
        entrada, // ENTRADA
        prazo, // PRAZO
        dataSaidaMaxima // SAÍDA MÁXIMA
      ];
      resposta.push(linhaAtual);
    }
  }
  if (resposta.length === 0) {
  resposta.push(['S/P', 'S/P', 'S/P', 'S/P']);
  console.log(resposta)
  }
  return resposta
}

function nada (){
  return 0;
}
//linha_relatorio('JUR II')

function pendencia(aba) {
    var prazos = filtrarLinhasNumericas(aba, 'd');
    var dataAtual = getDataDeHoje();
    var pend = [];
    for (var i = 0; i < prazos.length; i++) {
        var entrada = pegador(`A${prazos[i]}`, aba);
        if ((pegador(`P${prazos[i]}`, aba) != 'PROVIDENCIADO') && (subtrairDatas(entrada, dataAtual) >= 2)) {
            console.log('Verificando Pendências')
            var linhaPendencia = `${prazos[i]}`;
            pend.push(linhaPendencia);
        }
    }
    // Verifica se o vetor de pendências está vazio
    if (pend.length === 0) {
        return 'Não detectado';
    } else {
        return `Detectado: ${pend}`
    }
}



//ÁREA PARA GERAR OS RELATÓRIOS IMPRIMÍVEIS
function gerar_Relatorio() {
  //dados = [['Nome', 'Idade'],['Gustavo','20'],['Andrey','19']]
  jur1 = linha_relatorio('JUR I')
  jur2 = linha_relatorio('JUR II')
  jur3 = linha_relatorio('JUR III')
  jur4 = linha_relatorio('JUR IV')
  //Browser.msgBox("Relatório", mensagem, Browser.Buttons.OK);
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutput(gerarPaginaHTML(jur1,jur2,jur3,jur4)).setHeight(3508).setWidth(2480);
  ui.showModalDialog(html,'RELATÓRIO')
}

function mes (data) {
  console.log('Função Mês: '+data)
  meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
  return meses[data-1]
  
}

function gerarPaginaHTML(jur1,jur2,jur3,jur4) {
  var html = `
<!DOCTYPE html>
<html lang="pt-br">
  <head>
    <meta charset="UTF-8">
   <meta name="viewport" content="width=device-width, initial-scale=1.0">
   <meta name="author" content="Gustavo Maciel">
   <meta name="description" content="Sofware de levantamentos estatísticos">
   <title>Sistema de Estatísticas Documentais - Apoio Estatístico e Tecnológico da Conjur</title>
    <style type="text/css">
      * {
        margin: 0;
        padding: 0;
       box-sizing: border-box;
      }
      body {
      user-select: text;
      }
      header {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 1rem;

        border-bottom: 1px solid #000;

        padding: 2rem 0;
      }

      header img {
        width: 3.5rem;
        height: 4rem;
      }

      header ul {
        text-transform: uppercase;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-direction: column;
        gap: 0.25rem;
        list-style: none;
      }

      header .conjur {
        font-weight: bold;
      }
      .tabela_chegada {
        width: 100%;
      }
      .left-text {
      text-align: left;
    }

    .right-text {
      text-align: right;
    }
    .fixarRodape {
        bottom: 0;
        position: fixed;
        width: 100%;
        text-align: center;
    }
  </style>
  </head>
  <body>
  <header>
    <img src="https://upload.wikimedia.org/wikipedia/commons/b/bc/Bras%C3%A3o_do_Par%C3%A1.svg" alt="">
    <ul>
      <li>Governo do Estado do Pará</li>
      <li>Secretaria de Estado de Segurança Pública e defesa social</li>
      <li>Polícia Militar do Pará</li>
      <li class="conjur">Consultoria Jurídica</li>
    </ul>
    <img src="https://upload.wikimedia.org/wikipedia/commons/e/e3/Bras%C3%A3o_PMPA.PNG" alt="">
    </header>
    <form>
      <input id="botaoImprimir" type="button" value="IMPRIMIR RELATÓRIO" onClick="imprimirRelatorio();"/>
    </form>
    <table style="width: 100%;">
      <tr>
        <td style="text-align: left;">P1 - Apoio Estatístico e Tecnológico</td>
        <td style="text-align: right;">Belém/PA, ${getDataDeHoje('dia')} de ${mes(getDataDeHoje('mes'))} de ${getDataDeHoje('ano')}</td>
      </tr>
    </table>
    </br>
    <div class="chegadas">
      <table border="1" class="tabela_chegada">
      <caption>RELAÇÃO DE CHEGADAS DE PROCESOS</caption>
        <tr>
            <th style="width: 180px;">JURÍDICO</th>
            <th>ULTIMA CHEGADA</th>
            <th>PENDÊNCIAS</th>
        </tr>
        <tr>
            <td>JUR I</td>
            <td>${data_chegada('JUR I', 'a')}</td>
            <td>${pendencia('JUR I')}</td><!---  Mudar pra Função de Detecção de Pendências--->
        </tr>
        <tr>
            <td>JUR II</td>
            <td>${data_chegada('JUR II', 'a')}</td>
            <td>${pendencia('JUR II')}</td><!---  Mudar pra Função de Detecção de Pendências--->
        </tr>
        <tr>
            <td>JUR III</td>
            <td>${data_chegada('JUR III', 'a')}</td>
            <td>${pendencia('JUR III')}</td><!---  Mudar pra Função de Detecção de Pendências--->
        </tr>
        <tr>
            <td>JUR IV</td>
            <td>${data_chegada('JUR IV', 'a')}</td>
            <td>${pendencia('JUR IV')}</td><!---  Mudar pra Função de Detecção de Pendências--->
        </tr>
      </table>
      ${gerarTabelaHTML(jur1, 'JUR I')}
      ${gerarTabelaHTML(jur2, 'JUR II')}
      ${gerarTabelaHTML(jur3, 'JUR III')}
      ${gerarTabelaHTML(jur4, 'JUR IV')}
    </div>
    <script>
      function imprimirRelatorio() {
        document.getElementById("botaoImprimir").style.display = "none"; // Oculta o botão
        window.print();
        document.getElementById("botaoImprimir").style.display = "block"; //Habilita novamente
      }
    </script>
    <p><i><b>S/P</b> - Sem Pendências na planilha do protocolo</i></p>
    <footer class="fixarRodape">
    <hr>
    <p>Rod. </p>
    <p><b>Contato:  | E-mail: </b></p>
    <!--<p><i>Desenvolvido por: Gustavo Maciel | Git: @quietbytesilence </i></p>-->
    </footer>
  </body>
</html>
  `;

  return html;
}
function gerarTabelaHTML(parametros, titulo) {
  // Inicializa a string HTML da tabela
  var tabelaHTML = `<h4 style="text-align: center; width:100%"><b>${titulo}</b></h4>
                    <table border="1" class="tabela_chegada">
                      <tr>
                        <th style="width: 150px;">PAE:</th>
                        <th >ENTRADA</th>
                        <th style="width: 100px;">PRAZO</th>
                        <th>SAÍDA MÁX</th>
                      </tr>`;

  // Adiciona as linhas à tabela com base nos parâmetros
  for (var i = 0; i < parametros.length; i++) {
    tabelaHTML += '<tr>';
    // Adiciona cada item do parâmetro como uma célula na linha
    for (var j = 0; j < parametros[i].length; j++) {
      if (j == 2 && (parametros[i][j] != 'S/P')) {center = 'style="text-align: center;"'; dias = ' DIAS'}else{center=''; dias=''}
      if (parametros[i][j] == 'S/P') {center = 'style="text-align: center;"'}
      tabelaHTML += '<td ' + center + '>' + parametros[i][j] + dias + '</td>';
    }
    tabelaHTML += '</tr>';
  }

  // Fecha a tabela
  tabelaHTML += '</table>';

  return tabelaHTML;
}





