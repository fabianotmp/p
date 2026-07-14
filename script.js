function listarImagens() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página2");
  planilha.clear();
  const idsDasPastas = [""];
  const linhasParaInserir = []; 
  for (const id of idsDasPastas) {
    const arquivos = DriveApp.getFolderById(id).getFiles();
    while (arquivos.hasNext()) {
      const item = arquivos.next();
      const nome = item.getName().replace(/\.png$/i, "");
      const link = `https://lh3.googleusercontent.com/d/${item.getId()}=s341`;
      linhasParaInserir.push([nome, link]);
    }
  }
  if (linhasParaInserir.length > 0) {
    planilha.getRange(2, 1, linhasParaInserir.length, 2).setValues(linhasParaInserir);
  }
}

function listarPDF() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página2");
  planilha.clear();
  const idsDasPastas = [""];
  const linhasParaInserir = [];
  for (const id of idsDasPastas) {
    const arquivos = DriveApp.getFolderById(id).getFiles();
    while (arquivos.hasNext()) {
      const item = arquivos.next();
      const nome = item.getName().replace(/\.pdf$/i, "");
      const link = `https://drive.google.com/file/d/${item.getId()}/view`;
      const capa = `https://lh3.googleusercontent.com/d/${item.getId()}=s350`;
      linhasParaInserir.push([nome, link, capa]);
    }
  }
  if (linhasParaInserir.length > 0) {
    planilha.getRange(2, 1, linhasParaInserir.length, 3).setValues(linhasParaInserir);
  }
}

function listarVideos() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página2");
  planilha.clear();
  const idsDasPastas = [""];
  const linhasParaInserir = [];
  for (const id of idsDasPastas) {
    const arquivos = DriveApp.getFolderById(id).getFiles();
    while (arquivos.hasNext()) {
      const item = arquivos.next();
      const nome = item.getName().replace(/\.mp4$/i, "");
      const link = `https://drive.google.com/file/d/${item.getId()}/preview`;
      const thumbnail = `https://drive.google.com/thumbnail?id=${item.getId()}&sz=w400`;
      linhasParaInserir.push([nome, link, thumbnail]);
    }
  }
  if (linhasParaInserir.length > 0) {
    planilha.getRange(2, 1, linhasParaInserir.length, 3).setValues(linhasParaInserir);
  }
}

function listarAudios() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página2");
  planilha.clear();
  const idsDasPastas = [""];
  const linhasParaInserir = [];
  for (const id of idsDasPastas) {
    if (!id) continue;
    const arquivos = DriveApp.getFolderById(id).getFiles();
    while (arquivos.hasNext()) {
      const item = arquivos.next();
      const nome = item.getName().replace(/\.mp3$/i, "");
      const link = `https://drive.google.com/file/d/${item.getId()}/preview`;
      linhasParaInserir.push([nome, link]);
    }
  }
  if (linhasParaInserir.length > 0) {
    planilha.getRange(2, 1, linhasParaInserir.length, 2).setValues(linhasParaInserir);
  }
}

function listarJogos() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página2");
  planilha.clear();
  const idsDasPastas = [""];
  const linhasParaInserir = [];
  for (const id of idsDasPastas) {
    const arquivos = DriveApp.getFolderById(id).getFiles();  
    while (arquivos.hasNext()) {
      const item = arquivos.next();
      const nome = item.getName().replace(/\.zip$/i, "");
      const link = `https://drive.google.com/uc?export=download&id=${item.getId()}`;
      const youtube = `https://www.youtube.com/results?search_query=${encodeURIComponent(nome + " Game Boy Advance")}`;   
      linhasParaInserir.push([nome, link, youtube]);
    }
  }
  if (linhasParaInserir.length > 0) {
    planilha.getRange(2, 1, linhasParaInserir.length, 3).setValues(linhasParaInserir);
  }
}

function listarPlaylist() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página2");
  planilha.clear();
  const playlistId = ""; 
  const linhasParaInserir = [];
  let nextPageToken = "";
  try {
    do {
      const resultado = YouTube.PlaylistItems.list("snippet", {
        playlistId: playlistId,
        maxResults: 50,
        pageToken: nextPageToken
      });
      for (const item of resultado.items) {
        const videoId = item.snippet.resourceId.videoId;
        const nome = item.snippet.title;
        const link = `https://www.youtube.com/watch?v=${videoId}?rel=0`;
        linhasParaInserir.push([nome, link]);
      }
      nextPageToken = resultado.nextPageToken;
    } while (nextPageToken);
    if (linhasParaInserir.length > 0) {
      planilha.getRange(2, 1, linhasParaInserir.length, 2).setValues(linhasParaInserir);
    }
  } catch (erro) {
    Logger.log(erro.toString());
  }
}