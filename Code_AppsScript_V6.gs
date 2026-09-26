const CONFIG = {
  SPREADSHEET_ID: '1nBRbVwEFKOoNXkl1T3Pugj-doLjKwXgHRhOfoTykd34',
  INPUT_SHEET: 'Entrada_Fotos',
  CONFIG_SHEET: 'Config_OCR',
  PHOTO_FOLDER: 'Fotos_Entregas_OCR',
  OCR_LANGUAGE: 'pt',
  DEFAULT_CITY: 'Ourilândia do Norte',
  DEFAULT_UF: 'PA',
  INITIAL_STATUS: 'A conferir'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestão de Entregas Coutinho Logística')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processImage(payload) {
  requireAuth_(payload && payload.authToken);
  if (!payload || !payload.dataUrl) {
    throw new Error('Nenhuma foto foi recebida.');
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.INPUT_SHEET);
  if (!sheet) throw new Error('Aba Entrada_Fotos não encontrada.');

  const config = readConfig_(ss);
  const image = dataUrlToBlob_(payload.dataUrl);
  const folder = getOrCreateFolder_(config.photoFolder);

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const safeCode = String(payload.packageCode || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
  image.setName('ETIQUETA_' + stamp + (safeCode ? '_' + safeCode : '') + '.jpg');

  const photoFile = folder.createFile(image);
  const ocrText = runOcr_(image.copyBlob(), config.ocrLanguage);
  const parsed = parseLabel_(ocrText, config.defaultCity, config.defaultUf);

  const row = firstEmptyRow_(sheet, 3, 2, 500); // coluna C, linhas 2:501
  const now = new Date();
  const platform = String(payload.platform || 'Outro');
  const packageCode = String(payload.packageCode || parsed.packageCode || '');

  // B:M. A contém fórmula de ID na planilha.
  sheet.getRange(row, 2, 1, 12).setValues([[
    now,                                  // B Data/Hora
    photoFile.getUrl(),                   // C Foto
    platform,                             // D Plataforma
    packageCode,                          // E Código pacote
    ocrText,                              // F Texto OCR
    parsed.address,                       // G Endereço
    parsed.cep,                           // H CEP
    parsed.city || config.defaultCity,    // I Cidade
    parsed.uf || config.defaultUf,        // J UF
    parsed.confidence,                    // K Confiança
    config.initialStatus,                 // L Status OCR
    parsed.note                           // M Observação
  ]]);

  // Preenche automaticamente Bairro, Destinatario e Telefone na conferencia.
  const conf = ss.getSheetByName('Conferencia_OCR');
  if (conf) {
    conf.getRange(row, 9, 1, 3).setValues([[parsed.bairro || '', parsed.recipient || '', parsed.phone || '']]);
  }

  SpreadsheetApp.flush();
  const id = sheet.getRange(row, 1).getDisplayValue() || ('FOTO-' + Utilities.formatString('%04d', row - 1));
  const addressParts = splitStreetNumber_(parsed.address || '');

  return {
    ok: true,
    id: id,
    row: row,
    packageCode: packageCode,
    address: parsed.address,
    street: addressParts.street || '',
    number: addressParts.number || '',
    bairro: parsed.bairro || '',
    recipient: parsed.recipient || '',
    phone: parsed.phone || '',
    cep: parsed.cep,
    city: parsed.city || config.defaultCity,
    uf: parsed.uf || config.defaultUf,
    confidence: parsed.confidence,
    photoUrl: photoFile.getUrl()
  };
}

// V5.1 - Aprovação diretamente pelo aplicativo móvel.
function approveFromMobile(payload) {
  requireAuth_(payload && payload.authToken);
  if (!payload || !payload.row || !payload.id) throw new Error('Registro OCR inválido para aprovação.');

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const conf = ss.getSheetByName('Conferencia_OCR');
  if (!conf) throw new Error('Aba Conferencia_OCR não encontrada.');

  const row = Number(payload.row);
  if (!Number.isInteger(row) || row < 2) throw new Error('Linha de conferência inválida.');

  const currentId = String(conf.getRange(row, 1).getDisplayValue() || '').trim();
  if (currentId !== String(payload.id || '').trim()) {
    throw new Error('O registro mudou na planilha. Faça uma nova leitura da etiqueta.');
  }

  const street = String(payload.street || '').trim();
  const number = String(payload.number || '').trim();
  const bairro = String(payload.bairro || '').trim();
  const recipient = String(payload.recipient || '').trim();
  const phone = String(payload.phone || '').trim();
  const city = String(payload.city || '').trim();
  const uf = String(payload.uf || '').trim().toUpperCase();
  const cep = String(payload.cep || '').trim();
  const entregador = String(payload.entregador || '').trim();

  if (!street) throw new Error('Informe a rua/logradouro antes de aprovar.');
  if (!number) throw new Error('Informe o número do endereço antes de aprovar.');

  const enderecoFinal = (street + ' ' + number).trim();

  // H:M = Endereço final, Bairro, Destinatário, Telefone, Entregador, Aprovado?
  conf.getRange(row, 8, 1, 6).setValues([[
    enderecoFinal, bairro, recipient, phone, entregador, 'SIM'
  ]]);
  if (cep) conf.getRange(row, 5).setValue(cep);
  if (city) conf.getRange(row, 6).setValue(city);
  if (uf) conf.getRange(row, 7).setValue(uf);

  // Alterações feitas por script não disparam onEdit; transfere explicitamente.
  transferApprovedRow_(ss, row);
  SpreadsheetApp.flush();

  return {
    ok: true,
    id: payload.id,
    address: enderecoFinal,
    bairro: bairro,
    recipient: recipient,
    message: 'Entrega aprovada e adicionada à Base_Entregas.'
  };
}

function readConfig_(ss) {
  const sh = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  const out = {
    photoFolder: CONFIG.PHOTO_FOLDER,
    ocrLanguage: CONFIG.OCR_LANGUAGE,
    defaultCity: CONFIG.DEFAULT_CITY,
    defaultUf: CONFIG.DEFAULT_UF,
    initialStatus: CONFIG.INITIAL_STATUS
  };
  if (!sh) return out;

  const values = sh.getRange(3, 1, Math.max(sh.getLastRow() - 2, 1), 2).getDisplayValues();
  values.forEach(function(r) {
    const k = String(r[0] || '').trim();
    const v = String(r[1] || '').trim();
    if (k === 'Pasta para fotos' && v) out.photoFolder = v;
    if (k === 'Idioma OCR' && v) out.ocrLanguage = v;
    if (k === 'Cidade padrão' && v) out.defaultCity = v;
    if (k === 'UF padrão' && v) out.defaultUf = v;
    if (k === 'Status inicial OCR' && v) out.initialStatus = v;
  });
  return out;
}

function dataUrlToBlob_(dataUrl) {
  const m = String(dataUrl).match(/^data:(image\/[A-Za-z0-9.+-]+);base64,(.+)$/);
  if (!m) throw new Error('Formato de imagem inválido.');
  const mime = m[1];
  const bytes = Utilities.base64Decode(m[2]);
  return Utilities.newBlob(bytes, mime, 'etiqueta');
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function runOcr_(blob, language) {
  // Requer o serviço avançado "Drive API" ativado no Apps Script.
  // Editor do Apps Script > Serviços (+) > Drive API > Adicionar.
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error('Ative o serviço avançado Drive API no Apps Script: Serviços (+) > Drive API > Adicionar.');
  }

  let docFile;
  try {
    const resource = {
      name: 'OCR_TEMP_' + new Date().getTime(),
      mimeType: MimeType.GOOGLE_DOCS
    };

    // Apps Script Advanced Drive service v3. Ao converter imagem para Google Docs,
    // o Drive executa OCR automaticamente; ocrLanguage funciona como dica de idioma.
    docFile = Drive.Files.create(resource, blob, {
      ocrLanguage: language || 'pt',
      fields: 'id,name'
    });

    Utilities.sleep(1000);
    const text = DocumentApp.openById(docFile.id).getBody().getText();
    return String(text || '').trim();
  } catch (err) {
    throw new Error('Falha no OCR. Confirme que o serviço Drive API está ativado. Detalhe: ' + err.message);
  } finally {
    if (docFile && docFile.id) {
      try { Drive.Files.remove(docFile.id); } catch (e) {}
    }
  }
}

function parseLabel_(text, defaultCity, defaultUf) {
  const clean = String(text || '').replace(/\r/g, '\n').replace(/\n{2,}/g, '\n').trim();
  const lines = clean.split('\n').map(function(s) { return s.trim(); }).filter(Boolean);

  const cepMatch = clean.match(/\b(\d{5})[-\s]?(\d{3})\b/);
  const cep = cepMatch ? (cepMatch[1] + '-' + cepMatch[2]) : '';

  const ufMatch = clean.match(/\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b/i);
  const uf = ufMatch ? ufMatch[1].toUpperCase() : defaultUf;

  const streetWords = /\b(RUA|R\.|AVENIDA|AV\.|TRAVESSA|TV\.|ALAMEDA|RODOVIA|ROD\.|ESTRADA|EST\.|PASSAGEM|PÇ\.|PRACA|PRAÇA|VILA|QUADRA|Q\.|LOTE|LT\.)\b/i;
  let addressLine = '';
  for (let i = 0; i < lines.length; i++) {
    if (streetWords.test(lines[i])) {
      addressLine = lines[i];
      if (i + 1 < lines.length && /\b(?:N[º°o]?\.?\s*)?\d+[A-Za-z\/-]?\b/i.test(lines[i + 1])) {
        addressLine += ', ' + lines[i + 1];
      }
      break;
    }
  }

  if (!addressLine) {
    const candidate = lines.find(function(line) {
      return /\d/.test(line) && !/^\s*\d{5}[-\s]?\d{3}\s*$/.test(line) && line.length > 8;
    });
    addressLine = candidate || '';
  }

  let city = defaultCity;
  if (cepMatch) {
    const idx = lines.findIndex(function(line) { return line.indexOf(cepMatch[0]) >= 0; });
    if (idx >= 0) {
      const near = lines.slice(Math.max(0, idx - 2), idx + 2).join(' ');
      if (/Ouril[aâ]ndia\s+do\s+Norte/i.test(near)) city = 'Ourilândia do Norte';
    }
  }

  const packagePatterns = [
    /\b(?:MLB|MEL|SHP|BR)[A-Z0-9-]{6,}\b/i,
    /\b[A-Z]{2}\d{9}[A-Z]{2}\b/i,
    /\b\d{10,18}\b/
  ];
  let packageCode = '';
  for (let p = 0; p < packagePatterns.length && !packageCode; p++) {
    const m = clean.match(packagePatterns[p]);
    if (m) packageCode = m[0];
  }

  // Bairro: primeiro tenta o trecho após a vírgula; se não existir, procura
  // nas linhas logo depois do endereço (formato comum das etiquetas de entrega).
  let bairro = '';
  if (addressLine) {
    const addrClean = addressLine.replace(/^Endere[cç]o\s*:\s*/i, '');
    const parts = addrClean.split(',').map(function(x){ return x.trim(); }).filter(Boolean);
    if (parts.length > 1) bairro = parts[1].replace(/\bCEP\b.*$/i, '').trim();

    if (!bairro) {
      let idxEndereco = -1;
      for (let bi = 0; bi < lines.length; bi++) {
        const ln = String(lines[bi] || '').trim();
        if (ln && (addressLine === ln || addressLine.indexOf(ln) === 0 || ln.indexOf(addressLine) === 0)) { idxEndereco = bi; break; }
      }
      if (idxEndereco >= 0) {
        for (let bi = idxEndereco + 1; bi <= Math.min(lines.length - 1, idxEndereco + 3); bi++) {
          const candBairro = String(lines[bi] || '').trim();
          if (!candBairro) continue;
          if (/\b\d{5}[-\s]?\d{3}\b/.test(candBairro)) continue;
          if (new RegExp(defaultCity.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i').test(candBairro)) continue;
          if (/\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b/i.test(candBairro) && candBairro.length < 5) continue;
          if (/^[A-Za-zÀ-ÿ0-9 .'-]{3,45}$/.test(candBairro)) { bairro = candBairro; break; }
        }
      }
    }
  }

  // Destinatário: prioriza marcadores explícitos e depois pontua linhas próximas ao endereço.
  // Isso funciona melhor em etiquetas Mercado Livre/Shopee, onde o nome pode ficar
  // acima do endereço e nem sempre vem acompanhado da palavra "Destinatário".
  const recipient = extractRecipient_(lines, addressLine, clean);

  const phoneMatch = clean.match(/(?:\+?55\s*)?(?:\(?\d{2}\)?\s*)?9?\d{4}[-\s]?\d{4}/);
  const phone = phoneMatch ? phoneMatch[0].trim() : '';

  let confidence = 'Baixa';
  if (addressLine && cep) confidence = 'Alta';
  else if (addressLine || cep) confidence = 'Média';

  const noteParts = [];
  if (!addressLine) noteParts.push('Endereço não identificado automaticamente');
  if (!cep) noteParts.push('CEP não identificado');
  noteParts.push('Conferência humana obrigatória');

  return {
    address: addressLine,
    cep: cep,
    city: city,
    uf: uf,
    packageCode: packageCode,
    bairro: bairro,
    recipient: recipient,
    phone: phone,
    confidence: confidence,
    note: noteParts.join(' | ')
  };
}

function firstEmptyRow_(sheet, column, startRow, maxRows) {
  const values = sheet.getRange(startRow, column, maxRows, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) return startRow + i;
  }
  throw new Error('Limite de ' + maxRows + ' registros atingido na aba Entrada_Fotos.');
}

// ===== V3.2 - Transferencia automatica apos conferencia =====
function onEdit(e) {
  try {
    if (!e || !e.range || !e.source) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== 'Conferencia_OCR') return;
    if (e.range.getColumn() !== 13 || e.range.getRow() < 2) return; // coluna M = Aprovado?
    if (String(e.value || '').toUpperCase() !== 'SIM') return;

    // Usa a planilha vinculada ao evento. Isso evita chamadas que exigem
    // autorizacao adicional dentro do gatilho simples onEdit.
    transferApprovedRow_(e.source, e.range.getRow());
  } catch (err) {
    console.error(err);
  }
}

function transferApprovedRow_(ss, row) {
  const src = ss.getSheetByName('Conferencia_OCR');
  const dst = ss.getSheetByName('Base_Entregas');
  if (!src || !dst) throw new Error('Abas Conferencia_OCR ou Base_Entregas nao encontradas.');

  // A:M = ID Foto, Codigo, Plataforma, Endereco OCR, CEP, Cidade, UF,
  // Rua/Logradouro corrigido, Bairro, Destinatario, Telefone, Entregador, Aprovado?
  const v = src.getRange(row, 1, 1, 13).getDisplayValues()[0];
  const idFoto = String(v[0] || '').trim();
  const codigo = String(v[1] || '').trim();
  const plataforma = String(v[2] || '').trim();
  const enderecoFinal = String(v[7] || '').trim();
  const bairro = String(v[8] || '').trim();
  const destinatario = String(v[9] || '').trim();
  const telefone = String(v[10] || '').trim();
  const entregador = String(v[11] || '').trim();
  const cep = String(v[4] || '').trim();
  const cidade = String(v[5] || '').trim();
  const uf = String(v[6] || '').trim();

  if (!idFoto) throw new Error('ID Foto vazio na linha aprovada.');
  if (!enderecoFinal) throw new Error('Preencha o logradouro/endereco final antes de aprovar.');

  // Evita duplicidade pelo ID Origem OCR (coluna S) ou codigo do pacote (coluna A).
  const last = Math.max(dst.getLastRow(), 2);
  const existing = dst.getRange(2, 1, last - 1, 19).getDisplayValues();
  for (let i = 0; i < existing.length; i++) {
    if (String(existing[i][18] || '').trim() === idFoto) return;
    if (codigo && String(existing[i][0] || '').trim() === codigo) return;
  }

  const parsed = splitStreetNumber_(enderecoFinal);
  const nextRow = firstEmptyBaseRow_(dst);
  const idPacote = codigo || idFoto;

  dst.getRange(nextRow, 1, 1, 17).setValues([[
    idPacote,              // A ID Pacote
    plataforma,            // B Plataforma
    new Date(),            // C Data
    destinatario,          // D Destinatario
    telefone,              // E Telefone
    parsed.street,         // F Endereco/Rua
    parsed.number,         // G Numero
    '',                    // H Complemento
    bairro,                // I Bairro
    cidade,                // J Cidade
    uf,                    // K UF
    cep,                   // L CEP
    entregador,            // M Entregador
    '',                    // N Sequencia
    'Pendente',            // O Status
    0,                     // P Tentativas
    'Origem OCR - ' + idFoto // Q Ocorrencia
  ]]);

  // R = mapa por rua/endereco; S = ID origem OCR.
  const mapFormula = '=HYPERLINK("https://www.google.com/maps/search/?api=1&query="&ENCODEURL(F' + nextRow + '&IF(G' + nextRow + '<>"";", "&G' + nextRow + ';"")&IF(I' + nextRow + '<>"";", "&I' + nextRow + ';"")&IF(J' + nextRow + '<>"";", "&J' + nextRow + ';"")&IF(K' + nextRow + '<>"";" - "&K' + nextRow + ';"")&IF(L' + nextRow + '<>"";", CEP "&L' + nextRow + ';""));"Abrir no Maps")';
  dst.getRange(nextRow, 18).setFormula(mapFormula);
  dst.getRange(nextRow, 19).setValue(idFoto);

  // Marca o OCR como conferido na Entrada_Fotos quando encontrar o ID.
  const input = ss.getSheetByName('Entrada_Fotos');
  if (input) {
    const ids = input.getRange(2, 1, Math.max(input.getLastRow() - 1, 1), 1).getDisplayValues();
    for (let r = 0; r < ids.length; r++) {
      if (String(ids[r][0] || '').trim() === idFoto) {
        input.getRange(r + 2, 12).setValue('Conferido');
        break;
      }
    }
  }
  sincronizarAbas_(ss);
}

function splitStreetNumber_(text) {
  let s = String(text || '').trim();
  s = s.replace(/^Endere[cç]o\s*:\s*/i, '').replace(/\s*,?\s*CEP\s*:?\s*\d{5}-?\d{3}.*$/i, '').trim();
  // Remove bairro/cidade apos virgula para impedir que nomes de estabelecimentos ou referencias contaminem o Maps.
  const firstPart = s.split(',')[0].trim();
  const m = firstPart.match(/^(.*?)(?:\s*,?\s+)(\d+[A-Za-z0-9\/-]*)\s*$/);
  if (m) return { street: m[1].trim(), number: m[2].trim() };
  return { street: firstPart, number: '' };
}

function firstEmptyBaseRow_(sheet) {
  const startRow = 2;
  const maxRows = 500;
  const values = sheet.getRange(startRow, 19, maxRows, 1).getDisplayValues(); // coluna S
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) return startRow + i;
  }
  throw new Error('Limite de 500 entregas atingido na Base_Entregas.');
}

// ===== V4 - Estratificacao + Roteirizacao por proximidade =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ROTAS')
    .addItem('GERAR ROTA', 'gerarRotaMenu')
    .addItem('SINCRONIZAR APROVADOS', 'sincronizarAprovados')
    .addItem('SINCRONIZAR ABAS', 'sincronizarAbas')
    .addToUi();
}

function sincronizarAprovados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName('Conferencia_OCR');
  if (!src) throw new Error('Aba Conferencia_OCR nao encontrada.');

  const lastRow = src.getLastRow();
  let count = 0;
  for (let row = 2; row <= lastRow; row++) {
    const aprovado = String(src.getRange(row, 13).getDisplayValue() || '').trim().toUpperCase();
    if (aprovado === 'SIM') {
      const before = ss.getSheetByName('Base_Entregas').getLastRow();
      transferApprovedRow_(ss, row);
      const after = ss.getSheetByName('Base_Entregas').getLastRow();
      if (after > before) count++;
    }
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Sincronizacao concluida. Novas entregas transferidas: ' + count);
}

function gerarRotaMenu() {
  try {
    const r = gerarRotaFixa();
    SpreadsheetApp.getUi().alert(r.message + (r.rejected && r.rejected.length ? '\nEndereços não localizados: ' + r.rejected.length : ''));
  } catch (err) {
    SpreadsheetApp.getUi().alert('Erro ao gerar rota: ' + err.message);
  }
}

function gerarRota() {
  return gerarRotaFixa();
}

function gerarRotaFixa() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const cfg = ss.getSheetByName('Config_Rota');
  const base = ss.getSheetByName('Base_Entregas');
  const rota = ss.getSheetByName('Rota_Dia');
  if (!cfg || !base || !rota) throw new Error('Abas Config_Rota, Base_Entregas ou Rota_Dia não encontradas.');

  const startAddress = String(cfg.getRange('B3').getDisplayValue() || '').trim();
  const plusCode = String(cfg.getRange('B9').getDisplayValue() || '').trim();
  if (!startAddress) throw new Error('Ponto inicial fixo não informado em Config_Rota!B3.');

  let startGeo = geocodeAddress_(startAddress);
  if (!startGeo && plusCode) startGeo = geocodeAddress_(plusCode);
  if (!startGeo) throw new Error('Não foi possível localizar o ponto inicial fixo. Revise Config_Rota!B3 ou o Plus Code em B9.');

  cfg.getRange('A9:B12').setValues([
    ['Plus Code', plusCode],
    ['Latitude inicial fixa', startGeo.lat],
    ['Longitude inicial fixa', startGeo.lng],
    ['Status configuração', 'PRONTO - ponto fixo geocodificado']
  ]);
  cfg.getRange('B10:B11').setNumberFormat('0.000000');

  const driverFilter = String(cfg.getRange('B4').getDisplayValue() || '').trim();
  const statuses = String(cfg.getRange('B5').getDisplayValue() || 'Pendente,Em rota,Reagendada')
    .split(',').map(function(s){ return s.trim().toLowerCase(); }).filter(Boolean);

  const lastRow = base.getLastRow();
  if (lastRow < 2) throw new Error('Não há entregas na Base_Entregas.');
  const rows = base.getRange(2, 1, lastRow - 1, 23).getDisplayValues();

  const items = [];
  const rejected = [];
  rows.forEach(function(r, i) {
    const id = String(r[0] || '').trim();
    if (!id) return;
    const status = String(r[14] || '').trim();
    const driver = String(r[12] || '').trim();
    if (statuses.indexOf(status.toLowerCase()) < 0) return;
    if (driverFilter && driver.toLowerCase() !== driverFilter.toLowerCase()) return;

    const address = composeAddress_(r);
    if (!address) return;
    let itemLat = parseFloat(String(r[20] || '').replace(',', '.'));
    let itemLng = parseFloat(String(r[21] || '').replace(',', '.'));
    if (!isFinite(itemLat) || !isFinite(itemLng)) {
      const g = geocodeAddress_(address);
      if (!g) {
        rejected.push(id);
        return;
      }
      itemLat = g.lat;
      itemLng = g.lng;
      base.getRange(i + 2, 21, 1, 2).setValues([[itemLat, itemLng]]);
      Utilities.sleep(80);
    }

    items.push({
      baseRow: i + 2,
      id: id,
      destinatario: r[3],
      rua: r[5],
      numero: r[6],
      bairro: r[8],
      cidade: r[9],
      uf: r[10],
      cep: r[11],
      entregador: r[12],
      status: status,
      lat: itemLat,
      lng: itemLng
    });
  });

  if (!items.length) {
    rota.getRange(2, 1, Math.max(rota.getMaxRows() - 1, 1), 15).clearContent();
    throw new Error('Nenhuma entrega atende aos filtros de Config_Rota.');
  }

  const ordered = nearestNeighbor_(startGeo, items);
  ordered.forEach(function(item, idx) {
    base.getRange(item.baseRow, 14).setValue(idx + 1);
    base.getRange(item.baseRow, 23).setValue(item.distanceFromPrevious);
  });

  rota.getRange(2, 1, Math.max(rota.getMaxRows() - 1, 1), 15).clearContent();
  const out = ordered.map(function(item, idx) {
    return [
      idx + 1,
      item.id,
      item.destinatario,
      item.rua,
      item.numero,
      item.bairro,
      item.cidade,
      item.uf,
      item.cep,
      item.entregador,
      item.status,
      item.lat,
      item.lng,
      item.distanceFromPrevious,
      ''
    ];
  });
  rota.getRange(2, 1, out.length, 15).setValues(out);
  for (let i = 0; i < out.length; i++) {
    const row = i + 2;
    const url = 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(
      [out[i][3], out[i][4], out[i][5], out[i][6], out[i][7], out[i][8]].filter(Boolean).join(', ')
    );
    rota.getRange(row, 15).setFormula('=HYPERLINK("' + url + '";"Abrir no Maps")');
  }
  rota.getRange(2, 12, out.length, 2).setNumberFormat('0.000000');
  rota.getRange(2, 14, out.length, 1).setNumberFormat('0.00');
  sincronizarAbas_(ss);
  SpreadsheetApp.flush();

  return {
    ok: true,
    stops: ordered.length,
    rejected: rejected,
    startAddress: startAddress,
    startLat: startGeo.lat,
    startLng: startGeo.lng,
    route: ordered.map(function(item, idx) {
      return {
        sequence: idx + 1,
        id: item.id,
        recipient: item.destinatario || '',
        street: item.rua || '',
        number: item.numero || '',
        bairro: item.bairro || '',
        city: item.cidade || '',
        uf: item.uf || '',
        cep: item.cep || '',
        driver: item.entregador || '',
        status: item.status || 'Pendente',
        distance: Number(item.distanceFromPrevious || 0),
        mapUrl: 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(
          [item.rua, item.numero, item.bairro, item.cidade, item.uf, item.cep].filter(Boolean).join(', ')
        )
      };
    }),
    message: 'Rota gerada com ' + ordered.length + ' paradas a partir do ponto fixo: ' + startAddress
  };
}


// ===== V5.3 - Rotas públicas do aplicativo protegidas por sessão =====
function gerarRotaFixaMobile(authToken) {
  requireAuth_(authToken);
  return gerarRotaFixa();
}

function listarRotaMobile(authToken) {
  requireAuth_(authToken);
  return listarRotaMobile_();
}

// ===== V5.2 - Operacao da rota diretamente no aplicativo =====
function listarRotaMobile_( ) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const rota = ss.getSheetByName('Rota_Dia');
  const base = ss.getSheetByName('Base_Entregas');
  if (!rota || !base) throw new Error('Abas Rota_Dia ou Base_Entregas não encontradas.');

  const last = rota.getLastRow();
  if (last < 2) return {ok:true, total:0, pending:0, completed:0, route:[]};
  const rows = rota.getRange(2, 1, last - 1, 15).getDisplayValues();

  const baseLast = base.getLastRow();
  const baseRows = baseLast >= 2 ? base.getRange(2, 1, baseLast - 1, 23).getDisplayValues() : [];
  const statusById = {};
  baseRows.forEach(function(r){
    const id = String(r[0] || '').trim();
    if (id) statusById[id] = String(r[14] || '').trim();
  });

  const items = [];
  let completed = 0;
  rows.forEach(function(r) {
    const id = String(r[1] || '').trim();
    if (!id) return;
    const status = statusById[id] || String(r[10] || '').trim() || 'Pendente';
    if (status.toLowerCase() === 'entregue') { completed++; return; }
    const address = [r[3], r[4], r[5], r[6], r[7], r[8]].filter(Boolean).join(', ');
    items.push({
      sequence: Number(r[0]) || items.length + 1,
      id: id,
      recipient: r[2] || '',
      street: r[3] || '',
      number: r[4] || '',
      bairro: r[5] || '',
      city: r[6] || '',
      uf: r[7] || '',
      cep: r[8] || '',
      driver: r[9] || '',
      status: status,
      distance: Number(String(r[13] || '0').replace(',', '.')) || 0,
      mapUrl: 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(address)
    });
  });

  return {
    ok: true,
    total: items.length + completed,
    pending: items.length,
    completed: completed,
    route: items
  };
}

function concluirEntregaMobile(id, authToken) {
  requireAuth_(authToken);
  id = String(id || '').trim();
  if (!id) throw new Error('ID da entrega não informado.');

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const base = ss.getSheetByName('Base_Entregas');
  const rota = ss.getSheetByName('Rota_Dia');
  if (!base || !rota) throw new Error('Abas Base_Entregas ou Rota_Dia não encontradas.');

  let found = false;
  const last = base.getLastRow();
  if (last >= 2) {
    const ids = base.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() === id) {
        base.getRange(i + 2, 15).setValue('Entregue');
        found = true;
        break;
      }
    }
  }
  if (!found) throw new Error('Entrega não encontrada na Base_Entregas: ' + id);

  const rLast = rota.getLastRow();
  if (rLast >= 2) {
    const ids = rota.getRange(2, 2, rLast - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() === id) {
        rota.getRange(i + 2, 11).setValue('Entregue');
        break;
      }
    }
  }

  sincronizarAbas_(ss);
  SpreadsheetApp.flush();
  const updated = listarRotaMobile_();
  updated.message = 'Entrega ' + id + ' concluída com sucesso.';
  return updated;
}


// ===== V5.3 - Extração de destinatário melhorada =====
function extractRecipient_(lines, addressLine, fullText) {
  const explicitPatterns = [
    /^(?:DESTINAT[ÁA]RIO|DESTINATARIO|RECEBEDOR|CLIENTE|NOME)\s*[:\-]\s*(.+)$/i,
    /^(?:ENTREGAR\s+A|ENTREGA\s+PARA)\s*[:\-]?\s*(.+)$/i
  ];

  for (let i = 0; i < lines.length; i++) {
    for (let p = 0; p < explicitPatterns.length; p++) {
      const m = String(lines[i] || '').match(explicitPatterns[p]);
      if (m && isRecipientCandidate_(m[1])) return cleanRecipient_(m[1]);
    }
  }

  let addrIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    const ln = String(lines[i] || '').trim();
    if (!ln) continue;
    if (addressLine && (ln === addressLine || addressLine.indexOf(ln) === 0 || ln.indexOf(addressLine) === 0)) {
      addrIdx = i;
      break;
    }
  }

  const scored = [];
  const start = addrIdx >= 0 ? Math.max(0, addrIdx - 8) : 0;
  const end = addrIdx >= 0 ? Math.min(lines.length - 1, addrIdx + 2) : Math.min(lines.length - 1, 14);
  for (let i = start; i <= end; i++) {
    const raw = String(lines[i] || '').trim();
    if (!isRecipientCandidate_(raw)) continue;
    let score = 0;
    const distance = addrIdx >= 0 ? Math.abs(addrIdx - i) : 20;
    if (i < addrIdx) score += Math.max(0, 16 - distance * 2);
    if (/destinat|recebedor|cliente|entrega\s+para/i.test(String(lines[Math.max(0, i - 1)] || ''))) score += 20;
    const words = cleanRecipient_(raw).split(/\s+/).filter(Boolean).length;
    if (words >= 2 && words <= 5) score += 8;
    if (words >= 3 && words <= 4) score += 3;
    if (/^[A-ZÀ-Ý][A-Za-zÀ-ÿ'’-]+(?:\s+[A-ZÀ-Ý][A-Za-zÀ-ÿ'’-]+)+$/.test(cleanRecipient_(raw))) score += 3;
    scored.push({value: cleanRecipient_(raw), score: score});
  }
  scored.sort(function(a,b){ return b.score - a.score; });
  return scored.length ? scored[0].value : '';
}

function cleanRecipient_(value) {
  return String(value || '')
    .replace(/^(?:DESTINAT[ÁA]RIO|DESTINATARIO|RECEBEDOR|CLIENTE|NOME|ENTREGAR\s+A|ENTREGA\s+PARA)\s*[:\-]?\s*/i, '')
    .replace(/\s*\([A-Z0-9]{5,}\)\s*$/i, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function isRecipientCandidate_(value) {
  const cand = cleanRecipient_(value);
  if (!cand || cand.length < 5 || cand.length > 70) return false;
  if ((cand.match(/\d/g) || []).length > 1) return false;
  if (!/^[A-Za-zÀ-ÿ'’ .-]+$/.test(cand)) return false;
  const words = cand.split(/\s+/).filter(Boolean);
  if (words.length < 2 || words.length > 7) return false;
  if (/\b(MERCADO\s*LIVRE|SHOPEE|REMETENTE|VENDEDOR|ORIGEM|TRANSPORTADORA|DISTRIBUI[CÇ][AÃ]O|LOGISTICA|LOGÍSTICA|NOTA\s*FISCAL|VENDA|DESPACHAR|SKU|PEDIDO|PACOTE|ETIQUETA|ENDERE[CÇ]O|RUA|AVENIDA|TRAVESSA|CEP|BRASIL|PARQUE|CENTRO)\b/i.test(cand)) return false;
  return true;
}

// ===== V5.3 - Autenticação simples no servidor =====
const APP_AUTH = {
  DEFAULT_USER: 'clayton',
  DEFAULT_PASSWORD: 'Coutinho@2026',
  SESSION_SECONDS: 8 * 60 * 60
};

function sha256Hex_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''), Utilities.Charset.UTF_8);
  return bytes.map(function(b){ const v = b < 0 ? b + 256 : b; return ('0' + v.toString(16)).slice(-2); }).join('');
}

function getAuthCredentials_() {
  const p = PropertiesService.getScriptProperties();
  let user = p.getProperty('APP_USER');
  let hash = p.getProperty('APP_PASSWORD_HASH');
  if (!user || !hash) {
    user = APP_AUTH.DEFAULT_USER;
    hash = sha256Hex_(APP_AUTH.DEFAULT_PASSWORD);
    p.setProperties({APP_USER:user, APP_PASSWORD_HASH:hash}, false);
  }
  return {user:user, hash:hash};
}

function loginMobile(user, password) {
  user = String(user || '').trim();
  password = String(password || '');
  const cred = getAuthCredentials_();
  if (user.toLowerCase() !== cred.user.toLowerCase() || sha256Hex_(password) !== cred.hash) {
    throw new Error('Usuário ou senha inválidos.');
  }
  const token = Utilities.getUuid() + '-' + Utilities.getUuid();
  CacheService.getScriptCache().put('auth:' + token, cred.user, APP_AUTH.SESSION_SECONDS);
  return {ok:true, token:token, user:cred.user, expiresIn:APP_AUTH.SESSION_SECONDS};
}

function requireAuth_(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Sessão expirada. Faça login novamente.');
  const cache = CacheService.getScriptCache();
  const user = cache.get('auth:' + token);
  if (!user) throw new Error('Sessão expirada. Faça login novamente.');
  cache.put('auth:' + token, user, APP_AUTH.SESSION_SECONDS);
  return user;
}

function validarSessaoMobile(token) {
  const user = requireAuth_(token);
  return {ok:true, user:user};
}

function logoutMobile(token) {
  token = String(token || '').trim();
  if (token) CacheService.getScriptCache().remove('auth:' + token);
  return {ok:true};
}

function alterarSenhaMobile(authToken, senhaAtual, novaSenha) {
  const user = requireAuth_(authToken);
  novaSenha = String(novaSenha || '');
  if (novaSenha.length < 6) throw new Error('A nova senha deve ter pelo menos 6 caracteres.');
  const cred = getAuthCredentials_();
  if (sha256Hex_(String(senhaAtual || '')) !== cred.hash) throw new Error('Senha atual incorreta.');
  PropertiesService.getScriptProperties().setProperty('APP_PASSWORD_HASH', sha256Hex_(novaSenha));
  CacheService.getScriptCache().remove('auth:' + authToken);
  return {ok:true, user:user, message:'Senha alterada. Faça login novamente.'};
}

function sincronizarAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sincronizarAbas_(ss);
}

function sincronizarAbas_(ss) {
  const base = ss.getSheetByName('Base_Entregas');
  if (!base) throw new Error('Aba Base_Entregas nao encontrada.');
  const lastRow = base.getLastRow();
  const values = lastRow >= 2 ? base.getRange(2, 1, lastRow - 1, 23).getDisplayValues() : [];

  const ruas = [['ID Pacote','ID Origem OCR','Rua','Bairro','Cidade','Status']];
  const numeros = [['ID Pacote','ID Origem OCR','Numero','Rua','Bairro','Status']];
  const bairros = [['ID Pacote','ID Origem OCR','Bairro','Rua','Numero','Status']];
  const dests = [['ID Pacote','ID Origem OCR','Destinatario','Telefone','Rua','Numero','Bairro','Status']];

  values.forEach(function(r) {
    const id = String(r[0] || '').trim();
    if (!id) return;
    const origem = r[18] || '';
    const rua = r[5] || '';
    const numero = r[6] || '';
    const bairro = r[8] || '';
    const cidade = r[9] || '';
    const status = r[14] || '';
    ruas.push([id, origem, rua, bairro, cidade, status]);
    numeros.push([id, origem, numero, rua, bairro, status]);
    bairros.push([id, origem, bairro, rua, numero, status]);
    dests.push([id, origem, r[3] || '', r[4] || '', rua, numero, bairro, status]);
  });

  replaceSheetBody_(ss.getSheetByName('Ruas'), ruas);
  replaceSheetBody_(ss.getSheetByName('Numeros'), numeros);
  replaceSheetBody_(ss.getSheetByName('Bairros'), bairros);
  replaceSheetBody_(ss.getSheetByName('Destinatarios'), dests);
}

function replaceSheetBody_(sheet, matrix) {
  if (!sheet) return;
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
  if (matrix.length > 1) {
    sheet.getRange(2, 1, matrix.length - 1, matrix[0].length).setValues(matrix.slice(1));
  }
}

function composeAddress_(r) {
  return [r[5], r[6], r[8], r[9], r[10], r[11]].filter(function(x){ return String(x || '').trim(); }).join(', ');
}

function geocodeAddress_(address) {
  const response = Maps.newGeocoder().setLanguage('pt-BR').setRegion('br').geocode(address);
  if (!response || response.status !== 'OK' || !response.results || !response.results.length) return null;
  const loc = response.results[0].geometry.location;
  return {lat: loc.lat, lng: loc.lng};
}

function nearestNeighbor_(start, items) {
  const remaining = items.slice();
  const ordered = [];
  let current = {lat:start.lat, lng:start.lng};
  while (remaining.length) {
    let bestIndex = 0;
    let bestDistance = Infinity;
    for (let i = 0; i < remaining.length; i++) {
      const d = haversineKm_(current.lat, current.lng, remaining[i].lat, remaining[i].lng);
      if (d < bestDistance) {
        bestDistance = d;
        bestIndex = i;
      }
    }
    const next = remaining.splice(bestIndex, 1)[0];
    next.distanceFromPrevious = Math.round(bestDistance * 100) / 100;
    ordered.push(next);
    current = next;
  }
  return ordered;
}

function haversineKm_(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const toRad = function(v){ return v * Math.PI / 180; };
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
            Math.sin(dLon/2) * Math.sin(dLon/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}
