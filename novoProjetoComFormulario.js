// Formulário e planilha de respostas
var respostasFormulario = SpreadsheetApp.openById("Cole o ID aqui").getSheetByName('Nome da Planilha');
var formularioProjeto = FormApp.openById("Cole o ID aqui");
var ultimaResposta = respostasFormulario.getLastRow();
var datahoraResposta = respostasFormulario.getRange(`A${ultimaResposta}`).getValues();

// Crie um modelo no Google Docs e destaque os campos para serem substituídos

// Modelo de documento
var modeloProcuracao = driveapp.getfilebyid("");
var pastaDOC = driveapp.getfolderbyid("");

// Campos presentes no formulário e no modelo de documento
var nome = respostasFormulario.getRange(`B${ultimaResposta}`).getValues();
var cpf = respostasFormulario.getRange(`D${ultimaResposta}`).getValues();
var emailCliente = respostasFormulario.getRange(`X${ultimaResposta}`).getValues();

// Sugestão de formatação de data
var hoje = new Date();
var dd = String(hoje.getDate()).padStart(2, '0');
var mm = String(hoje.getMonth() + 1).padStart(2, '0'); //January is 0!
var yyyy = hoje.getFullYear();

hoje = dd + '/' + mm + '/' + yyyy;

// Criar um novo documento com dados do formulário
var formatoNome = nome + ' - ' + cpf + ' - ' +  hoje
var copiaModelo = modeloProcuracao.makeCopy(formatoNome, pastaDOC);
const novaDocumento = DocumentApp.openById(copiaModelo.getId());

// Função que povoa o novo documento

function preenchimentoAutomaticoProcuracao() {
  
  //Texto no arquivo DOC
  var textoDocumento = novaProcuracao.getBody();

  textoDocumento.replaceText("{{nome}}", nome);
  textoDocumento.replaceText("{{cpf}}", cpf);

  novaProcuracao.saveAndClose();

  //pastaDOC.removeFile(copiaModelo);
  console.log("Novo Doc criado! ");

  // Registro de IDs
  respostasFormulario.getRange(`Y${ultimaResposta}`).setValue(novaProcuracao.getUrl());
}

// Função para criar PDFs
function criarPDF() {

  const pastaPDF = DriveApp.getFolderById("");
  const pastaTemporaria = DriveApp.getFolderById("");
  const modeloDocumento = DriveApp.getFileById("");

  const novoArquivoTemporario = modeloDocumento.makeCopy(pastaTemporaria).setName(placa + "-" + servico + "-" + datahoraResposta + "-arquivo-temporario");
  const arquivoTemporario = DocumentApp.openById(novoArquivoTemporario.getId());
  
  console.log("Arquivo Temporário aberto");
  const textoArquivoTemporario = arquivoTemporario.getBody();

  textoArquivoTemporario.replaceText("{{nome}}", nome);
  textoArquivoTemporario.replaceText("{{cpf}}", cpf);

  arquivoTemporario.saveAndClose();

  console.log("Arquivo Temporário salvo")

  const blobNovoDocumentoPDF = arquivoTemporario.getAs(MimeType.PDF);
  const arquivoPDF =  pastaPDF.createFile(blobNovoDocumentoPDF).setName(formatoNome + ".pdf");
  pastaTemporaria.removeFile(novoArquivoTemporario);
  
  console.log("Arquivo PDF foi criado!");

  respostasFormulario.getRange(`Z${ultimaResposta}`).setValue(arquivoPDF.getId());
  console.log("ID registrado");
  
  return arquivoPDF

}

// Função que é usada após a submissão de formulário
function aposSubmissaoFormulario() {

  const arquivoPDFCriado = criarPDF();
  sendEmail(arquivoPDFCriado);

}

// Função de envio de e-mail com anexo
function sendEmail() {
  
  const linkDOC = respostasFormulario.getRange(`Y${ultimaResposta}`).getValue();
  const anexoPDF = DriveApp.getFileById(respostasFormulario.getRange(`Z${ultimaResposta}`).getValue());
  var destinatarios = `email1@hotmail.com,email2@gmail.com,${emailCliente}`;
  
  MailApp.sendEmail(
    {to: destinatarios,

    //Assunto:
    subject: nome + " - " + cpf + " - Documento super importante",

    //Mensagem:
    body: `Os arquivo PDF criado está em anexo.\nCaso seja necessário fazer alterações no texto manualmente, acesse o mesmo pelo link:\n${linkDOC}    
    \nMensagem automática.`,

    //Anexos: 
    attachments: [anexoPDF],
    name: 'REMETENTE'

    }
  );

  console.log("E-mail enviado com sucesso!")
}
