/**
 * Создает меню в таблице при открытии
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📍 Геокодер')
      .addItem('Найти координаты для пустых строк', 'massGeocode')
      .addToUi();
}

/**
 * Основная функция массового геокодирования
 */
function massGeocode() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Укажите номера столбцов (A=1, B=2, C=3 и т.д.)
  var colAddress = 2;   // Столбец B: Адрес
  var colIndex = 3;     // Столбец C: Индекс
  var colResult = 6;    // Столбец F: Сюда пишем Координаты (поменяйте на ваш)
  
  var apiKey = "КЛЮЧ АПИ";
  
  // Берем данные со 2-й строки до конца
  var range = sheet.getRange(2, 1, lastRow - 1, colResult);
  var data = range.getValues();
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var currentAddress = row[colAddress - 1];
    var currentIndex = row[colIndex - 1];
    var currentResult = row[colResult - 1];
    
    // Пропускаем, если адрес пуст или координаты уже найдены
    if (!currentAddress || (currentResult && currentResult.toString().indexOf(",") !== -1)) continue;
    
    // Формируем полный запрос: Индекс (если есть), Москва, Адрес
    var fullQuery = "Россия, Москва, " + (currentIndex ? currentIndex + ", " : "") + currentAddress;
    
    var url = "https://geocode-maps.yandex.ru/1.x/?apikey=" + apiKey + 
              "&geocode=" + encodeURIComponent(fullQuery) + 
              "&format=json&results=1";
              
    try {
      var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      if (response.getResponseCode() == 200) {
        var json = JSON.parse(response.getContentText());
        var members = json.response.GeoObjectCollection.featureMember;
        
        if (members.length > 0) {
          var pos = members[0].GeoObject.Point.pos.split(" ");
          var formattedCoords = pos[1] + ", " + pos[0];
          // Записываем результат сразу в ячейку
          sheet.getRange(i + 2, colResult).setValue(formattedCoords);
        } else {
          sheet.getRange(i + 2, colResult).setValue("Не найдено");
        }
      } else {
        sheet.getRange(i + 2, colResult).setValue("Ошибка API: " + response.getResponseCode());
      }
    } catch (e) {
      sheet.getRange(i + 2, colResult).setValue("Ошибка сети");
    }
    
    // ПАУЗА 200 мс (чтобы не превысить лимит запросов в секунду)
    Utilities.sleep(200);
  }
}