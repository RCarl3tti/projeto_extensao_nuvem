function cadastrarNovoPaciente() {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var urlArquivo = "";


  var nome = planilha.getRange('C6').getValue();
  var nome_mae = planilha.getRange('J6').getValue();
  var documento = planilha.getRange('C9').getValue();
  var data_nasc = planilha.getRange('F9').getValue();
  var idade = planilha.getRange('F12').getValue();
  var contato = planilha.getRange('C12').getValue();
  var contato_emergencia = planilha.getRange('C15').getValue();
  var particular_convenio = planilha.getRange('F15').getValue();
  var endereco = planilha.getRange('J9').getValue();
  var escolaridade = planilha.getRange('J12').getValue();
  var profissao = planilha.getRange('N12').getValue();
  var estado_civil = planilha.getRange('J15').getValue();
  var demanda = planilha.getRange('C18').getValue();
  var evolucao = planilha.getRange('J18').getValue();

  // Enviar para outra tabela com ID
  const outraTabela = planilha.getSheetByName("Lista de Pacientes"); 
  const ultimaLinha = outraTabela.getLastRow();
  const colunaId = 1; // Coluna onde o ID será inserido


  // Encontrar o maior ID existente e gerar um novo
  const idsExistentes = outraTabela.getRange(2, colunaId, ultimaLinha - 1).getValues().flat(); // Ignora a primeira linha (cabeçalho)
  const maiorId = Math.max(...idsExistentes, 0); // 0 como valor padrão se não houver IDs
  const novoId = maiorId + 1;

  // ID do documento modelo (substitua pelo ID real do seu modelo)
  var modeloId = '11X0qA3DAfGPB6p-LoU2sgoELv8p5Nrq8FhzBabB30N0'; // Exemplo: '1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890'
  
  // ID da pasta onde os novos documentos serão salvos
  var pastaId = '1K2dALG8cTnH1-FeXI4ZIKCdVbUH-T142'; // Substitua pelo ID da pasta onde deseja salvar o arquivo

  // Abrir o modelo de documento
  var modeloDoc = DriveApp.getFileById(modeloId);
  
  // Fazer uma cópia do modelo
  var novoDoc = modeloDoc.makeCopy(nome, DriveApp.getFolderById(pastaId));

  // Abrir o novo documento para edição
  var doc = DocumentApp.openById(novoDoc.getId());
  var body = doc.getBody();
  
  // Substituir os placeholders pelos valores reais
  body.replaceText('{{novoId}}', novoId);
  body.replaceText('{{nome}}', nome);
  body.replaceText('{{nome_mae}}', nome_mae);
  body.replaceText('{{documento}}', documento);
  body.replaceText('{{data_nasc}}', data_nasc);
  body.replaceText('{{idade}}', idade);
  body.replaceText('{{contato}}', contato);
  body.replaceText('{{contato_emergencia}}', contato_emergencia);
  body.replaceText('{{particular_convenio}}', particular_convenio);
  body.replaceText('{{endereco}}', endereco);
  body.replaceText('{{escolaridade}}', escolaridade);
  body.replaceText('{{profissao}}', profissao);
  body.replaceText('{{estado_civil}}', estado_civil);
  body.replaceText('{{demanda}}', demanda);
  body.replaceText('{{evolucao}}', evolucao);

  
  // Salvar e fechar o documento
  doc.saveAndClose();

  // Obter a URL do documento criado
  urlArquivo = doc.getUrl();


  // Inserir os dados com o novo ID e o link do arquivo de resumo e evolução
  const novaLinha = ultimaLinha + 1;
  const dadosComId = [novoId, nome, documento,data_nasc, contato, contato_emergencia, nome_mae, endereco, escolaridade, profissao, estado_civil, idade, particular_convenio, urlArquivo]; // Adiciona o ID no início e o link ao final do array de valores
  outraTabela.getRange(novaLinha, 1, 1, dadosComId.length).setValues([dadosComId]);



  SpreadsheetApp.getUi().alert("Cadastro realizado com sucesso!");

}
