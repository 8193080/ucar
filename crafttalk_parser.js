const FORM = FormApp.openById('1lVHeGMdRhuVFj2SJ7363EJnqfV3XY9kfrdq3RSLcHSY');
const MONITOR = SpreadsheetApp.openById('1APPmYvrBy--GVz0RW4U42YRqvz9Ce3F3zMVaNmhVmPg').getSheetByName('Монитор');
const PIVOTDATA = SpreadsheetApp.openById('1APPmYvrBy--GVz0RW4U42YRqvz9Ce3F3zMVaNmhVmPg').getSheetByName('Данные Сводная');
const STATSDATA = SpreadsheetApp.openById('1APPmYvrBy--GVz0RW4U42YRqvz9Ce3F3zMVaNmhVmPg').getSheetByName('Данные Онлайн');
const DIALOGDATA = SpreadsheetApp.openById('1APPmYvrBy--GVz0RW4U42YRqvz9Ce3F3zMVaNmhVmPg').getSheetByName('Данные Диалоги');
const LIST = SpreadsheetApp.openById('1APPmYvrBy--GVz0RW4U42YRqvz9Ce3F3zMVaNmhVmPg').getSheetByName('Обработчик');

//const YASHIFTS = SpreadsheetApp.openById('1upzRtdWrLT9dykag6HBvBOe88KPEHQQKt12b5r57zUM').getSheetByName('Данные');
//const LOCATIONS = SpreadsheetApp.openById('1upzRtdWrLT9dykag6HBvBOe88KPEHQQKt12b5r57zUM').getSheetByName('Списки');
//const REPORT = SpreadsheetApp.openById('1mNB5OzeHMKv1vtSy6YB9dPev8qU6FjEm8sI7jSGOjAk');

//функция для очистки массива от пустых значений
function clearArray(ar) {
  return (ar != null && ar != "" || ar === 0);
}

function parsing() {
  var queue = MONITOR.getRange('C2').getValue()
  if(MONITOR.getRange('B5').getValue() == "⌛ Идет обновление" || MONITOR.getRange('B5').getValue() == '') {
    MONITOR.getRange('C2').setValue(queue+1)
  }
  else {
    if( queue > 0 ) {
      MONITOR.getRange('C2').setValue(queue-1)
    }
    MONITOR.insertRowBefore(5)
    MONITOR.getRange('B5').setValue("⌛ Идет обновление")
    var responses = FORM.getResponses();
    var fileId = responses[responses.length-1-queue].getItemResponses()[0].getResponse()
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob().setContentType(MimeType.MICROSOFT_EXCEL)
      let resource = {
      title: blob.getName().split('.')[0],
      mimeType: MimeType.GOOGLE_SHEETS
    };
    let newfileId = Drive.Files.insert(resource, blob, {"convert":"true"}).getId();

    MONITOR.getRange('D5').setValue(file.getUrl())

    var fileStats = JSON.stringify(SpreadsheetApp.openById(newfileId).getSheets()[0].getRange('A2:T2').getValues()[0])
    var fileDialog = JSON.stringify(SpreadsheetApp.openById(newfileId).getSheets()[0].getRange('A1:AX1').getValues()[0])
    var trueStats = JSON.stringify(["Дата",	"ID Оператора",	"ФИО",	"Макс. одновременных диалогов",	"Оффлайн",	"В ожидании",	"На диалоге",	"В работе",	"В сети",	"Обед",	"Тренинг",	"Невидимый",	"Оффлайн(сек)",	"В ожидании(сек)",	"На диалоге(сек)",	"В работе(сек)",	"В сети(сек)",	"Обед(сек)",	"Тренинг(сек)",	"Невидимый(сек)"]);

    var truePivot = JSON.stringify(["Дата",	"ID Оператора",	"ФИО",	"Назначено сессий",	"Отвеченных сессий",	"Всего сообщений",	"Всего ответов",	"Среднее чистое время реакции",	"Среднее время закрытия",	"Среднее время первого ответа",	"Среднее время ответа",	"Средняя длина сообщения",	"Потерянных сессий",	"% потерянных сессий",	"Средняя клиентская оценка",	"Количество полученных оценок",	"Время ожидания клиентов оператором",	"Среднее время обслуживания",	"Параллельность чатов метод 1",	"Параллельность чатов метод 2"]);

    var trueDialog = JSON.stringify(["id диалога",	"id сессии",	"id оператора",	"ФИО оператора",	"Название агента",	"Время поступления обращения в веб чат/МП",	"Время старта сессии",	"Канал поступления",	"Идентификатор канала",	"Название канала",	"Тип канала",	"Тип клиента",	"id клиента",	"ФИО клиента",	"Телефон",	"E-mail",	"Результат соединения",	"Время ожидания в очереди (сек)",	"Результат ответа оператора",	"Количество Hold",	"Hold, сек",	"Время реакции на обращение (сек)",	"Время ожидания первого открытия диалога оператором",	"Время поступления обращения оператору",	"Длительность диалога (сек)",	"Длительность диалога с оператором (сек)",	"Время обработки, сек (AHT)",	"Количество итераций",	"Среднее время итераций, сек",	"Время завершения обработки обращения (сессии)",	"Причина завершения сессии",	"Время завершения обработки обращения (диалога)",	"Причина завершения диалога",	"Не завершен",	"Оценка",	"Тематика 1 уровень",	"Тематика 2 уровень",	"Тематика 3 уровень",	"Тематика 4 уровень",	"Тематика 5 уровень",	"Статья для ответа оператором",	"Статья, используемая оператором",	"Статья для ответа ботом",	"Классификатор статьи для ответа оператором",	"Классификатор статьи, используемой оператором",	"Классификатор для статьи бота",	"Группа",	"Язык",	"Навык",	"Проблема решена"])

    let time = new Date()

    if (fileStats == truePivot) {
      MONITOR.getRange('C5').setValue("Сводная")
      try {
        insertDataPivot(newfileId)
      } catch {MONITOR.getRange('B5').setValue("❌ Завершено с ошибкой" + time)}
    }    
    else if (fileStats == trueStats) {
      MONITOR.getRange('C5').setValue("Статистика")
      try {
        insertDataStats(newfileId)
      } catch {MONITOR.getRange('B5').setValue("❌ Завершено с ошибкой" + time)}
    }
    else if (fileDialog == trueDialog) {
      MONITOR.getRange('C5').setValue("Диалоги")
      try {
        insertDataDialog(newfileId)
      } catch {MONITOR.getRange('B5').setValue("❌ Завершено с ошибкой" + time)}
    }
    else {
      MONITOR.getRange('C5').setValue("Был загружен неверный файл (" + file.getName() + ")")
      MONITOR.getRange('B5').setValue("❌ Пропущено " + time)
    }
  }
}

function insertDataPivot(newfileId) {
  let NEWFILE = SpreadsheetApp.openById(newfileId).getSheets()[0]; //загружаемый файл

  let colNfId = NEWFILE.getRange('A:B').getDisplayValues().filter(clearArray); //даты и время выходов
  let colYsId = PIVOTDATA.getRange('A:B').getDisplayValues().filter(clearArray); //даты и время выходов из основного файла

  let newData = NEWFILE.getRange('A:T').getDisplayValues().filter(clearArray); //массив данных из загружаемого файла

  let rows = PIVOTDATA.getRange('A:A').getDisplayValues().filter(clearArray).length+1
  let rowsNF = NEWFILE.getRange('A:A').getDisplayValues().filter(clearArray).length+1

  for(var i = 1; i < rowsNF; i++) {

    let index = colYsId.some(function(item) {
      return JSON.stringify(item) === JSON.stringify(colNfId[i]);
    });

    if (index == false && colNfId[i][1].includes('@')) {
      PIVOTDATA.getRange(rows,1).setValue(newData[i][0])
      PIVOTDATA.getRange(rows,2).setValue(newData[i][1])
      PIVOTDATA.getRange(rows,3).setValue(newData[i][2])
      PIVOTDATA.getRange(rows,4).setValue(newData[i][3])
      PIVOTDATA.getRange(rows,5).setValue(newData[i][4])
      PIVOTDATA.getRange(rows,6).setValue(newData[i][5])
      PIVOTDATA.getRange(rows,7).setValue(newData[i][6])
      PIVOTDATA.getRange(rows,8).setValue(newData[i][7])
      PIVOTDATA.getRange(rows,9).setValue(newData[i][8])
      PIVOTDATA.getRange(rows,10).setValue(newData[i][9])
      PIVOTDATA.getRange(rows,11).setValue(newData[i][10])
      PIVOTDATA.getRange(rows,12).setValue(newData[i][11])
      PIVOTDATA.getRange(rows,13).setValue(newData[i][12])
      PIVOTDATA.getRange(rows,14).setValue(newData[i][13])
      PIVOTDATA.getRange(rows,15).setValue(newData[i][14])
      PIVOTDATA.getRange(rows,16).setValue(newData[i][15])
      PIVOTDATA.getRange(rows,17).setValue(newData[i][16])
      PIVOTDATA.getRange(rows,18).setValue(newData[i][17])
      PIVOTDATA.getRange(rows,19).setValue(newData[i][18])
      PIVOTDATA.getRange(rows,20).setValue(newData[i][19])
      PIVOTDATA.getRange(rows,21).setValue(newData[i][20])
      rows = rows + 1
    }
    else if (index == true && colNfId[i][1].includes('@')) {
      var position = 1
      for (var p = 0; p < rows; p++) {
        if (colYsId[p][0] == colNfId[i][0] && colYsId[p][1] == colNfId[i][1]) {
          position = p+1
          Logger.log(position)
          PIVOTDATA.getRange(position,1).setValue(newData[i][0])
          PIVOTDATA.getRange(position,2).setValue(newData[i][1])
          PIVOTDATA.getRange(position,3).setValue(newData[i][2])
          PIVOTDATA.getRange(position,4).setValue(newData[i][3])
          PIVOTDATA.getRange(position,5).setValue(newData[i][4])
          PIVOTDATA.getRange(position,6).setValue(newData[i][5])
          PIVOTDATA.getRange(position,7).setValue(newData[i][6])
          PIVOTDATA.getRange(position,8).setValue(newData[i][7])
          PIVOTDATA.getRange(position,9).setValue(newData[i][8])
          PIVOTDATA.getRange(position,10).setValue(newData[i][9])
          PIVOTDATA.getRange(position,11).setValue(newData[i][10])
          PIVOTDATA.getRange(position,12).setValue(newData[i][11])
          PIVOTDATA.getRange(position,13).setValue(newData[i][12])
          PIVOTDATA.getRange(position,14).setValue(newData[i][13])
          PIVOTDATA.getRange(position,15).setValue(newData[i][14])
          PIVOTDATA.getRange(position,16).setValue(newData[i][15])
          PIVOTDATA.getRange(position,17).setValue(newData[i][16])
          PIVOTDATA.getRange(position,18).setValue(newData[i][17])
          PIVOTDATA.getRange(position,19).setValue(newData[i][18])
          PIVOTDATA.getRange(position,20).setValue(newData[i][19])
          PIVOTDATA.getRange(position,21).setValue(newData[i][20])
        }   
      }
    }
  }
  let time = new Date()
  MONITOR.getRange('B5').setValue("✅ Обновлено " + time)

  Drive.Files.remove(newfileId)
  if (MONITOR.getRange('C2').getValue() > 0) {
    parsing()
  }
}

function insertDataStats(newfileId) {
  let NEWFILE = SpreadsheetApp.openById(newfileId).getSheets()[0]; //загружаемый файл
  let colNfId = NEWFILE.getRange('A:B').getDisplayValues().filter(clearArray); //даты и время выходов + почты
  let colYsId = STATSDATA.getRange('A:B').getDisplayValues().filter(clearArray); //даты и время выходов из основного файла
  let newData = NEWFILE.getRange('A:T').getDisplayValues().filter(clearArray); //массив данных из загружаемого файла

  let rowsNF = NEWFILE.getRange('A:A').getDisplayValues().filter(clearArray).length

  let arrNewData = []

  for(let i = 1; i < rowsNF; i++) {

    let index = colYsId.some(function(item) {
      return JSON.stringify(item) === JSON.stringify(colNfId[i]);
    });

    if (index == false  && newData[i][0].indexOf('-') == -1) {

      let arrNewDataRow = []

      arrNewDataRow.push(
        newData[i][0]
      , newData[i][1]
      , newData[i][2]
      , newData[i][3]
      , newData[i][4]
      , newData[i][5]
      , newData[i][6]
      , newData[i][7]
      , newData[i][8]
      , newData[i][9]
      , newData[i][10]
      , newData[i][11]
      , newData[i][12]
      , newData[i][13]
      , newData[i][14]
      , newData[i][15]
      , newData[i][16]
      , newData[i][17]
      , newData[i][18]
      , newData[i][19]
      , newData[i][20])

      arrNewData.push(arrNewDataRow)
    }
  }

  try {
    let lnght = arrNewData.length
    let arrAA = STATSDATA.getRange('A:A').getDisplayValues()
    let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
    let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1

    STATSDATA.insertRowsAfter(STATSDATA.getMaxRows(), lnght);
    STATSDATA.getRange(lastRow, 1, lnght, 21).setValues(arrNewData) 
  } catch {}
  

  var time = new Date()
  MONITOR.getRange('B5').setValue("✅ Обновлено " + time)

  Drive.Files.remove(newfileId)
  if (MONITOR.getRange('C2').getValue() > 0) {
    parsing()
  }
}

function insertDataDialog(newfileId) {
  var NEWFILE = SpreadsheetApp.openById(newfileId).getSheets()[0]; //загружаемый файл
  var colNfId = NEWFILE.getRange('B:B').getDisplayValues().filter(clearArray).flat(); //id session
  var colYsId = DIALOGDATA.getRange('B:B').getDisplayValues().filter(clearArray).flat(); //id session из основного файла
  var newData = NEWFILE.getRange('A:AX').getDisplayValues().filter(clearArray); //массив данных из загружаемого файла

  let arrNewData = []

  for(let i = 1; i < colNfId.length; i++) {

    let indexId = colYsId.indexOf(colNfId[i])

    if (indexId == -1 && newData[i][33] != 'Да' && (newData[i][3] != 'Принят ботом' || newData[i][30] == 'Закрыто по таймауту')) { //проверяем, что сессия завершена и отсутствует в таблице

      let arrNewDataRow = []

      arrNewDataRow.push(
        newData[i][0] //id диалога
      , newData[i][1] //id сессии
      , newData[i][2] //id оператора
      , newData[i][3] //ФИО оператора
      , newData[i][5] //Время поступления обращения в веб чат/МП
      , newData[i][6] //Время старта сессии
      , newData[i][12] //id клиента
      , newData[i][17] //Время ожидания в очереди (сек)
      , newData[i][18] //Результат ответа оператора
      , newData[i][19] //Количество Hold
      , newData[i][20] //Hold, сек
      , newData[i][21] //Время реакции на обращение (сек)
      , newData[i][22] //Время ожидания первого открытия диалога оператором
      , newData[i][23] //Время поступления обращения оператору
      , newData[i][24] //Длительность диалога (сек)
      , newData[i][25] //Длительность диалога с оператором (сек)
      , newData[i][26] //Время обработки, сек (AHT)
      , newData[i][27] //Количество итераций
      , newData[i][28] //Среднее время итераций, сек
      , newData[i][29] //Время завершения обработки обращения (сессии)
      , newData[i][30] //Причина завершения сессии
      , newData[i][31] //Время завершения обработки обращения (диалога)
      , newData[i][32] //Причина завершения диалога 
      , "" //пустой столбец
      , newData[i][34]) //Оценка

      arrNewData.push(arrNewDataRow)
    }
  }
  try {
    let lnght = arrNewData.length
    let arrAA = DIALOGDATA.getRange('A:A').getDisplayValues()
    let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
    let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1
    DIALOGDATA.insertRowsAfter(DIALOGDATA.getMaxRows(), lnght);
    DIALOGDATA.getRange(lastRow, 1, lnght, 25).setValues(arrNewData) 
  } catch {}

  var time = new Date()
  MONITOR.getRange('B5').setValue("✅ Обновлено " + time)

  Drive.Files.remove(newfileId)

  if (MONITOR.getRange('C2').getValue() > 0) {
    parsing()
  }
}

function role (qst) {
  var ans
  var locList = LIST.getRange('N2:P100').getValues()  
    for (let i = 0; i < locList.length; i++) {
    if (qst.includes(locList[i][0])) {
      ans = locList[i][1]
      break
    }
  }
  return (ans)
}
