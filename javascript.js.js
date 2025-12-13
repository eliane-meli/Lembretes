// Google Apps Script para integração com a planilha
function doGet(e) {
  const action = e.parameter.action;
  
  try {
    const sheet = SpreadsheetApp.openById('1k33SVvjo-gjoZG1cWVrvm48QjUrPjtsSuEnGIYmHcUQ')
      .getSheetByName('Atividades diárias');
    
    if (!sheet) {
      return createResponse(false, 'Planilha não encontrada');
    }
    
    if (action === 'getActivities') {
      return getActivitiesData(sheet);
    } else if (action === 'complete') {
      return markActivityComplete(e.parameter, sheet);
    }
    
    return createResponse(false, 'Ação não reconhecida');
    
  } catch (error) {
    return createResponse(false, error.toString());
  }
}

function getActivitiesData(sheet) {
  const data = sheet.getDataRange().getValues();
  const activities = [];
  
  // Pular cabeçalho (linha 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] && row[0].toString().trim() !== '') {
      activities.push({
        name: row[0].toString().trim(),
        time: row[1] ? formatTime(row[1]) : '00:00',
        rowIndex: i + 1
      });
    }
  }
  
  return createResponse(true, {
    activities: activities,
    total: activities.length,
    sheetName: sheet.getName(),
    timestamp: new Date().toISOString()
  });
}

function markActivityComplete(params, sheet) {
  const activityName = params.activity;
  const activityTime = params.time;
  const rowIndex = params.rowIndex || 'auto';
  const timestamp = params.timestamp || new Date().toISOString();
  
  let rowToUpdate = -1;
  
  if (rowIndex !== 'auto') {
    rowToUpdate = parseInt(rowIndex);
  } else {
    // Buscar a linha automaticamente
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] === activityName && formatTime(row[1]) === activityTime) {
        rowToUpdate = i + 1;
        break;
      }
    }
  }
  
  if (rowToUpdate > 0 && rowToUpdate <= sheet.getLastRow()) {
    // Atualizar coluna C (índice 2) com "CONCLUÍDO"
    sheet.getRange(rowToUpdate, 3).setValue('CONCLUÍDO');
    // Atualizar coluna D (índice 3) com a data/hora
    sheet.getRange(rowToUpdate, 4).setValue(new Date(timestamp));
    
    return createResponse(true, {
      message: 'Atividade marcada como concluída',
      rowUpdated: rowToUpdate,
      activity: activityName,
      time: activityTime
    });
  } else {
    // Adicionar nova linha
    sheet.appendRow([activityName, activityTime, 'CONCLUÍDO', new Date(timestamp)]);
    
    return createResponse(true, {
      message: 'Atividade adicionada como concluída',
      activity: activityName,
      time: activityTime
    });
  }
}

function formatTime(timeValue) {
  if (!timeValue) return '00:00';
  
  if (timeValue instanceof Date) {
    const hours = timeValue.getHours().toString().padStart(2, '0');
    const minutes = timeValue.getMinutes().toString().padStart(2, '0');
    return `${hours}:${minutes}`;
  }
  
  if (typeof timeValue === 'string') {
    // Extrair horário HH:MM da string
    const timeMatch = timeValue.match(/(\d{1,2}):(\d{2})/);
    if (timeMatch) {
      const hours = timeMatch[1].padStart(2, '0');
      const minutes = timeMatch[2].padStart(2, '0');
      return `${hours}:${minutes}`;
    }
  }
  
  return timeValue.toString();
}

function createResponse(success, data) {
  const response = {
    success: success,
    ...(typeof data === 'string' ? { message: data } : data)
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*');
}

// Teste local
function testGetActivities() {
  const result = getActivitiesData(SpreadsheetApp.openById('1k33SVvjo-gjoZG1cWVrvm48QjUrPjtsSuEnGIYmHcUQ')
    .getSheetByName('Atividades diárias'));
  Logger.log(JSON.stringify(result, null, 2));
}
