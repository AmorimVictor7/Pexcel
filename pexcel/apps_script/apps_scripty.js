const sheetDestino = "Incluir Ação";
const sheetOrigem = "Planilha2";

const ss = SpreadsheetApp.getActiveSpreadsheet();


//  PEGAR DADOS PARA SELECTS (SEM DUPLICADOS)
function doGet() {

const sheet = ss.getSheetByName(sheetOrigem);

const dados = sheet.getDataRange().getValues();

dados.shift(); // remove cabeçalho

// usar Set para evitar duplicação
const resultado = {
tipo_acao: new Set(),
curso: new Set(),
campanha: new Set(),
marca: new Set(),
ativo: new Set()
};

dados.forEach(linha => {

// A → Ativo
if (linha[0])
resultado.ativo.add(linha[0]);

// E → Marca
if (linha[4])
resultado.marca.add(linha[4]);

// F → Curso
if (linha[5])
resultado.curso.add(linha[5]);

// H → Tipo de ação
if (linha[7])
resultado.tipo_acao.add(linha[7]);

// J → Campanha
if (linha[9])
resultado.campanha.add(linha[9]);

});


// converter Set → Array
const resposta = {
tipo_acao: [...resultado.tipo_acao],
curso: [...resultado.curso],
campanha: [...resultado.campanha],
marca: [...resultado.marca],
ativo: [...resultado.ativo]
};

return ContentService
.createTextOutput(JSON.stringify(resposta))
.setMimeType(ContentService.MimeType.JSON);

}



// SALVAR DADOS COM DATA FORMATADA
function doPost(e) {

try {

const sheet = ss.getSheetByName(sheetDestino);

const data = JSON.parse(e.postData.contents);


//  FORMATAR DATA dd/MM/yyyy
const dataFormatada = Utilities.formatDate(
new Date(data.data_acao),
Session.getScriptTimeZone(),
"dd/MM/yyyy"
);


// salvar linha
sheet.appendRow([

data.tipo_acao,   // Coluna A
data.curso,       // Coluna B
data.campanha,    // Coluna C
data.marca,       // Coluna D
dataFormatada,    // Coluna E (data formatada)
data.ativo,       // Coluna F
data.link         // Coluna G

]);

return ContentService
.createTextOutput(JSON.stringify({
status: "sucesso"
}))
.setMimeType(ContentService.MimeType.JSON);

}

catch(error) {

return ContentService
.createTextOutput(JSON.stringify({
status: "erro",
mensagem: error.toString()
}))
.setMimeType(ContentService.MimeType.JSON);

}

}