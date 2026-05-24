// =============================================================
//  CONECTA GENERAL BRUCE E ARREDORES
//  Código do Google Apps Script — versão atualizada
//
//  ESTRUTURA DA ABA "Cadastros" (ordem exata das colunas):
//    A: timestamp       B: nome            C: bloco
//    D: servico         E: categoria       F: descricao
//    G: whatsapp        H: autorizado      I: aprovado
//    J: imagem_url      K: entrega         L: area_entrega
//    M: taxa_entrega
//
//  As colunas K, L e M são NOVAS — acrescente-as no cabeçalho
//  da planilha antes de usar este código.
//
//  ESTRUTURA DA ABA "Solicitações" (criada automaticamente):
//    A: timestamp       B: status          C: nome_solicitante
//    D: whatsapp        E: anuncio         F: acao            G: mensagem
//
//  COMO ATUALIZAR O SCRIPT:
//  1. Acesse script.google.com e abra o projeto do site.
//  2. Selecione tudo (Ctrl+A) e apague.
//  3. Cole este código inteiro e salve (Ctrl+S).
//  4. Clique em Implantar → Gerenciar implantações → editar (lápis)
//     → Nova versão → Implantar.
//  5. A URL permanece a mesma. Nada muda no site.
// =============================================================

// ID da pasta do Google Drive onde as imagens são salvas.
// Encontre na URL da pasta: drive.google.com/drive/folders/SEU_ID_AQUI
// Deixe "" para desativar o salvamento de imagens.
var PASTA_DRIVE_ID = "";

// Nomes das abas — só altere se renomear as abas na planilha
var ABA_CADASTROS    = "Cadastros";
var ABA_SOLICITACOES = "Solicitações";


// =============================================================
//  GET — retorna lista de serviços APROVADOS para o site
// =============================================================
function doGet(e) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba      = planilha.getSheetByName(ABA_CADASTROS);

    if (!aba) {
      return jsonResponse({ ok: false, error: "Aba '" + ABA_CADASTROS + "' não encontrada." });
    }

    var dados = aba.getDataRange().getValues();
    if (dados.length <= 1) {
      return jsonResponse({ ok: true, servicos: [] });
    }

    var servicos = [];

    // i = 1 para pular o cabeçalho (linha 0)
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];

      // Coluna I (índice 8): "aprovado" — publica somente se for "SIM"
      var aprovado = String(linha[8]).trim().toUpperCase();
      if (aprovado !== "SIM") continue;

      servicos.push({
        nome:         String(linha[1]  || ""),   // B
        bloco:        String(linha[2]  || ""),   // C
        servico:      String(linha[3]  || ""),   // D
        categoria:    String(linha[4]  || ""),   // E
        descricao:    String(linha[5]  || ""),   // F
        whatsapp:     String(linha[6]  || ""),   // G
        imagem_url:   String(linha[9]  || ""),   // J
        entrega:      String(linha[10] || ""),   // K (novo)
        area_entrega: String(linha[11] || ""),   // L (novo)
        taxa_entrega: String(linha[12] || "")    // M (novo)
      });
    }

    return jsonResponse({ ok: true, servicos: servicos });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}


// =============================================================
//  POST — recebe novo cadastro ou solicitação de alteração/remoção
// =============================================================
function doPost(e) {
  try {
    var params = e.parameter || {};
    var tipo   = params.tipo_solicitacao || "novo_cadastro";

    if (tipo === "remocao_alteracao") {
      return processarSolicitacao(params);
    } else {
      return processarCadastro(params);
    }

  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}


// =============================================================
//  Processa NOVO CADASTRO → grava na aba "Cadastros" como PENDENTE
// =============================================================
function processarCadastro(p) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba      = planilha.getSheetByName(ABA_CADASTROS);

  if (!aba) {
    return jsonResponse({ ok: false, error: "Aba '" + ABA_CADASTROS + "' não encontrada." });
  }

  if (!p.nome || !p.servico || !p.whatsapp) {
    return jsonResponse({ ok: false, error: "Campos obrigatórios ausentes." });
  }

  if (p.autorizado !== "true") {
    return jsonResponse({ ok: false, error: "Autorização não concedida." });
  }

  // Salva imagem no Drive se houver
  var imagemUrl = "";
  if (p.imagem_base64 && PASTA_DRIVE_ID) {
    imagemUrl = salvarImagemDrive(p.imagem_base64, p.imagem_nome || "imagem.jpg");
  }

  // Grava linha respeitando a ordem das colunas da planilha existente
  // A coluna "aprovado" (I) começa VAZIA → a responsável preenche "SIM" para aprovar
  aba.appendRow([
    new Date(),              // A: timestamp
    limpar(p.nome),          // B: nome
    limpar(p.bloco),         // C: bloco / localização
    limpar(p.servico),       // D: serviço
    limpar(p.categoria),     // E: categoria
    limpar(p.descricao),     // F: descrição
    limpar(p.whatsapp),      // G: whatsapp
    "SIM",                   // H: autorizado (só chega aqui se autorizou)
    "",                      // I: aprovado (PENDENTE — preencha "SIM" para publicar)
    imagemUrl,               // J: imagem_url
    limpar(p.entrega),       // K: entrega (novo)
    limpar(p.area_entrega),  // L: area_entrega (novo)
    limpar(p.taxa_entrega)   // M: taxa_entrega (novo)
  ]);

  return jsonResponse({ ok: true, mensagem: "Cadastro recebido. Aguardando aprovação." });
}


// =============================================================
//  Processa SOLICITAÇÃO DE ALTERAÇÃO/REMOÇÃO → aba "Solicitações"
// =============================================================
function processarSolicitacao(p) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba      = planilha.getSheetByName(ABA_SOLICITACOES);

  // Cria a aba automaticamente se não existir ainda
  if (!aba) {
    aba = planilha.insertSheet(ABA_SOLICITACOES);
    aba.appendRow([
      "Data/Hora", "Status", "Nome do solicitante",
      "WhatsApp", "Anúncio referenciado",
      "Ação solicitada", "Mensagem"
    ]);
    // Formata o cabeçalho
    var cab = aba.getRange(1, 1, 1, 7);
    cab.setFontWeight("bold")
       .setBackground("#0f3a66")
       .setFontColor("#ffffff");
    aba.setFrozenRows(1);
  }

  if (!p.nome_solicitante || !p.whatsapp_solicitante) {
    return jsonResponse({ ok: false, error: "Campos obrigatórios ausentes na solicitação." });
  }

  aba.appendRow([
    new Date(),                          // A: timestamp
    "PENDENTE",                          // B: status (para controle interno)
    limpar(p.nome_solicitante),          // C: nome
    limpar(p.whatsapp_solicitante),      // D: whatsapp
    limpar(p.anuncio_referencia),        // E: anúncio referenciado
    limpar(p.acao_solicitada),           // F: ação (Alterar / Remover)
    limpar(p.mensagem_solicitacao)       // G: mensagem / detalhes
  ]);

  return jsonResponse({ ok: true, mensagem: "Solicitação recebida com sucesso." });
}


// =============================================================
//  Salva imagem base64 no Google Drive e retorna URL pública
// =============================================================
function salvarImagemDrive(base64, nomeArquivo) {
  try {
    var partes        = base64.split(",");
    var mime          = partes[0].match(/:(.*?);/)[1];
    var dadosBinarios = Utilities.base64Decode(partes[1]);
    var blob          = Utilities.newBlob(dadosBinarios, mime, nomeArquivo);

    var pasta   = DriveApp.getFolderById(PASTA_DRIVE_ID);
    var arquivo = pasta.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return "https://drive.google.com/uc?id=" + arquivo.getId();
  } catch (err) {
    Logger.log("Erro ao salvar imagem: " + err);
    return "";
  }
}


// =============================================================
//  Utilitários
// =============================================================

// Remove espaços e impede injeção de fórmulas na planilha
function limpar(valor) {
  if (!valor) return "";
  var v = String(valor).trim();
  if (v.startsWith("=") || v.startsWith("+") || v.startsWith("-") || v.startsWith("@")) {
    v = "'" + v;
  }
  return v;
}

// Retorna resposta JSON com MIME type correto
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
