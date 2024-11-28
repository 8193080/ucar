const FORM = FormApp.openById('1oOcL__p-8Ojy4Q3KJMaDJYxKVvwfK3jM17dNFsOfTc8');
const YASHIFTS = SpreadsheetApp.openById('1upzRtdWrLT9dykag6HBvBOe88KPEHQQKt12b5r57zUM').getSheetByName('Данные');
const VBSHIFTS = SpreadsheetApp.openById('1upzRtdWrLT9dykag6HBvBOe88KPEHQQKt12b5r57zUM').getSheetByName('Данные Велобайк');
const URENTSHIFTS = SpreadsheetApp.openById('1upzRtdWrLT9dykag6HBvBOe88KPEHQQKt12b5r57zUM').getSheetByName('Данные Юрент');
const LOCATIONS = SpreadsheetApp.openById('1upzRtdWrLT9dykag6HBvBOe88KPEHQQKt12b5r57zUM').getSheetByName('Списки');
const REPORT = SpreadsheetApp.openById('1mNB5OzeHMKv1vtSy6YB9dPev8qU6FjEm8sI7jSGOjAk');
const YAOPERATIONS = SpreadsheetApp.openById('1upzRtdWrLT9dykag6HBvBOe88KPEHQQKt12b5r57zUM').getSheetByName('Операции');
const MONITOR = SpreadsheetApp.openById('1upzRtdWrLT9dykag6HBvBOe88KPEHQQKt12b5r57zUM').getSheetByName('Монитор');


function clearArray(ar) {
  return (ar != null && ar != "" || ar === 0);
}

function parsing() {
  var t = MONITOR.getRange('A1').getValue()
  if(MONITOR.getRange('C5').getValue() == "⌛ Идет обновление") {
    MONITOR.getRange('A1').getValue()
  }
  else {
    t > 0 ? MONITOR.getRange('A1').setValue(t-1) :
    MONITOR.insertRowBefore(5)
    MONITOR.getRange('C5').setValue("⌛ Идет обновление")
    let responses = FORM.getResponses();
    let fileId = responses[responses.length-1-t].getItemResponses()[0].getResponse()
    let file = DriveApp.getFileById(fileId);
    let blob = file.getBlob().setContentType(MimeType.MICROSOFT_EXCEL)
      let resource = {
      title: blob.getName().split('.')[0],
      mimeType: MimeType.GOOGLE_SHEETS
    };
    let newfileId = Drive.Files.insert(resource, blob, {"convert":"true"}).getId();

    let fileTitle

    try {fileTitle = JSON.stringify(SpreadsheetApp.openById(newfileId).getSheetByName('Sheet1').getRange('A1:Q1').getValues()[0])} 
    catch {fileTitle = JSON.stringify(SpreadsheetApp.openById(newfileId).getSheetByName('export').getRange('A1:Q1').getValues()[0])}
    let trueOldTitle = JSON.stringify(["ID смены", "ID исполнителя", "ФИО исполнителя", "Статус смены", "Локация", "Специализация", "Тип транспорта", "Вместимость батарей", "Вместимость самокатов", "Дата (таймзона локации)", "Начало (план)", "Окончание (план)", "Начало (факт)", "Окончание (факт)", "Опоздание", "Ранний уход", "Невыход"])
    let trueNewTitle = JSON.stringify(["ID смены","ID исполнителя","ФИО исполнителя","Статус смены","Регион","Локация","Специализация","Тип транспорта","Вместимость батарей","Вместимость самокатов","Начало (план)","Окончание (план)","Начало (факт)","Окончание (факт)","Опоздание","Ранний уход","Невыход"])
    let operationTitle = JSON.stringify(["Дата","ID исполнителя","ID слота","ФИО исполнителя","Регион","Транспортное средство","Специализация","Вендор","Продолжительность смены","Своя машина","Завершённые миссии по замене аккумуляторов","Запланированные работы по замене аккумуляторов","Завершённые работы по замене аккумуляторов","Пропущенные работы по замене аккумуляторов","Завершённые миссии по отвозу на склад","Запланированные работы по отвозу на склад","Завершённые работы по отвозу на склад"])
    let velobikeDriversTitle = JSON.stringify(["ФИО исполнителя","Номер телефона исполнителя","Регион","Роль исполнителя","Начало (план)","Окончание (план)","Начало (факт)","Окончание (факт)","Количество операций (релокация)","Количество операций (сбор)","Количество операций (вывоз)","Количество операций (перезарядки)","","","","",""])
    let velobikeChargersTitle = JSON.stringify(["ФИО","Номер телефона","Регион","Роль исполнителя","Начало (план)","Окончание (план)","Начало (факт)","Окончание (факт)","Количество операций (перезарядки)","","","","","","","",""])
    let urentTitle = JSON.stringify(["ID смены","Город","Названия зон","Должности","ФИО исполнителя","Телефон исполнителя","ID исполнителя","ФИО руководителя","ID руководителя","ФИО постановщика","ID постановщика","Плановое время начала","Фактическое время начала","Плановое время окончания","Фактическое время окончания","Открытый комментарий","Закрытый комментарий"])

    if (fileTitle == trueOldTitle) {
      insertOldReport(newfileId)
    }
    else if (fileTitle == trueNewTitle) {
      MONITOR.getRange('D5').setValue('Выгрузка смен')
      MONITOR.getRange('E5').setValue(file.getUrl())
      insertNewReport(newfileId)
    }
    else if (fileTitle == operationTitle) {
      MONITOR.getRange('D5').setValue('Выгрузка операций')
      MONITOR.getRange('E5').setValue(file.getUrl())
      insertOperation(newfileId)
    }
    else if (fileTitle == velobikeDriversTitle) {
      MONITOR.getRange('D5').setValue('Выгрузка смен велобайк Водители')
      MONITOR.getRange('E5').setValue(file.getUrl())
      insertVelobikeShifts(newfileId, "drivers")
    }
    else if (fileTitle == velobikeChargersTitle) {
      MONITOR.getRange('D5').setValue('Выгрузка смен велобайк Чарджеры')
      MONITOR.getRange('E5').setValue(file.getUrl())
      insertVelobikeShifts(newfileId, "chargers")
    }
    else if (fileTitle == urentTitle) {
      MONITOR.getRange('D5').setValue('Выгрузка смен Юрент')
      MONITOR.getRange('E5').setValue(file.getUrl())
      insertUrentShifts(newfileId)
    }
    else {
      MONITOR.getRange('E5').setValue(file.getUrl())
      MONITOR.getRange('D5').setValue("❌ Загружен неверный файл (" + file.getName() + ")")
    }
  }
}


function insertUrentShifts(newfileId) {
  let NEWFILE = SpreadsheetApp.openById(newfileId).getSheetByName('export'); //загружаемый файл

  let colNfId = NEWFILE.getRange('A:A').getDisplayValues().filter(clearArray).flat(); //Ид смен из загружаемого файла
  let colYsId = URENTSHIFTS.getRange('A:A').getDisplayValues().filter(clearArray).flat(); //Ид смен из основного файла

  let newData = NEWFILE.getRange('A:U').getDisplayValues().filter(clearArray); //массив данных из загружаемого файла

  let arrNewData = []

  for(let i = 1; i < colNfId.length; i++) {

    let index = colYsId.indexOf(colNfId)

    if (index == -1) {

      let arrNewDataRow = []

      let owner = newData[i][4].includes("TD") ? "TD" : "Other"

      let role 
      
      if (newData[i][3].includes("Тестировщик")) {
        role = "Мастер приемщик"
      }
      else if (newData[i][3].includes("Скаут")) {
        role = "Скаут"
      }
      else {
        role = "Водитель"
      }

      let trueDate = new Date (Date.parse(newData[i][11]) + (7 * 60000 * 60))
      let time = new Date (trueDate).getHours()
      let type = time > 16 ? "Ночь" : "День"

      arrNewDataRow.push(
        newData[i][0] //Смена
      , newData[i][1] //Город
      , newData[i][3] //Роль
      , newData[i][4] //ФИО
      , newData[i][5] //Телефон
      , newData[i][6] //Исполнитель
      , new Date (Date.parse(newData[i][11]) + (7 * 60000 * 60)) //Старт план
      , new Date (Date.parse(newData[i][13]) + (7 * 60000 * 60)) //Окончание план
      , new Date (Date.parse(newData[i][12]) + (7 * 60000 * 60)) //Старт факт
      , new Date (Date.parse(newData[i][14]) + (7 * 60000 * 60)) //Окончание факт
      , newData[i][19] //Релоки
      , newData[i][18] //Сбор
      , newData[i][20] //Перезарядки
      , type //Тип смены
      , owner
      , role
      )

      arrNewData.push(arrNewDataRow)
    }
    else {
      Logger.log(newData[i][1])
    }
  }

  try {
    let lnght = arrNewData.length
    let arrAA = URENTSHIFTS.getRange('A:A').getDisplayValues()
    let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
    let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1

    URENTSHIFTS.insertRowsAfter(URENTSHIFTS.getMaxRows(), lnght);
    URENTSHIFTS.getRange(lastRow, 1, lnght, 16).setValues(arrNewData) 
  } 
  catch {
    console.log('nothing')
  }

  var time = new Date()
  MONITOR.getRange('C5').setValue("✅ Обновлено " + time)

  Drive.Files.remove(newfileId)
  
  if (MONITOR.getRange('A1').getValue() > 0) {
    parsing()
  }
}

function insertVelobikeShifts(newfileId, type) {
  let NEWFILE = SpreadsheetApp.openById(newfileId).getSheetByName('Sheet1'); //загружаемый файл

  let colNfId = NEWFILE.getRange('A:A').getDisplayValues().filter(clearArray).flat(); //Ид смен из загружаемого файла
  let colYsId = VBSHIFTS.getRange('N:N').getDisplayValues().filter(clearArray).flat(); //Ид смен из основного файла

  let newData = NEWFILE.getRange('A:L').getDisplayValues().filter(clearArray); //массив данных из загружаемого файла

  let arrNewData = []

  if (type == "drivers") {
    for(let i = 1; i < colNfId.length; i++) {

      let index = colYsId.indexOf(newData[i][1] + new Date (Date.parse(newData[i][6])))

      if (index == -1) {

        let arrNewDataRow = []

        let time = new Date (Date.parse(newData[i][4])).getHours()
        Logger.log(time)
        let type = time > 16 || time < 4 ? "Ночь" : "День"

        arrNewDataRow.push(
          newData[i][0] //ФИО
        , newData[i][1] //Телефон
        , newData[i][2] //Регион
        , newData[i][3] //Роль
        , new Date (Date.parse(newData[i][4])) //Старт план
        , new Date (Date.parse(newData[i][5])) //Окончание план
        , new Date (Date.parse(newData[i][6])) //Старт факт
        , new Date (Date.parse(newData[i][7])) //Окончание факт
        , newData[i][8] //Релоки
        , newData[i][9] //Сбор
        , newData[i][10] //Вывоз
        , newData[i][11] //Перезарядки
        , type //Тип смены
        , newData[i][1] + new Date (Date.parse(newData[i][6])) //Хэш
        )

        arrNewData.push(arrNewDataRow)
      }
      else {
        Logger.log(newData[i][1])
      }
    }
  }

  try {
    let lnght = arrNewData.length
    let arrAA = VBSHIFTS.getRange('A:A').getDisplayValues()
    let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
    let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1

    VBSHIFTS.insertRowsAfter(VBSHIFTS.getMaxRows(), lnght);
    VBSHIFTS.getRange(lastRow, 1, lnght, 14).setValues(arrNewData) 
  } 
  catch {
    console.log('nothing')
  }

  var time = new Date()
  MONITOR.getRange('C5').setValue("✅ Обновлено " + time)

  Drive.Files.remove(newfileId)

  if (MONITOR.getRange('A1').getValue() > 0) {
    parsing()
  }
}

function insertNewReport (newfileId) {
  let NEWFILE = SpreadsheetApp.openById(newfileId).getSheetByName('Sheet1'); //загружаемый файл

  let colNfId = NEWFILE.getRange('A:A').getDisplayValues().filter(clearArray).flat(); //Ид смен из загружаемого файла
  let colYsId = YASHIFTS.getRange('A:A').getDisplayValues().filter(clearArray).flat(); //Ид смен из основного файла

  let newData = NEWFILE.getRange('A:S').getDisplayValues().filter(clearArray); //массив данных из загружаемого файла

  let arrNewData = []

  for(let i = 1; i < colNfId.length; i++) {

    let index = colYsId.indexOf(colNfId[i])

    if (index == -1) {

      let arrNewDataRow = []

      arrNewDataRow.push(
        newData[i][0] //0-1
      , newData[i][1] //1-2
      , newData[i][2] //2-3
      , newData[i][3] //3-4
      , newData[i][5] //4-5
      , newData[i][6] //5-6
      , newData[i][7] //6-7
      , newData[i][8] //7-8
      , newData[i][9] //8-9
      , new Date (Date.parse(newData[i][10])).toLocaleDateString('ru-RU') //9-10
      , (newData[i][10].split(" "))[1] //10-11 начало план
      , (newData[i][11].split(" "))[1] //11-12
      , (newData[i][12].split(" "))[1] //12-13
      , (newData[i][13].split(" "))[1] //13-14 окончание факт
      , newData[i][14] //14-15
      , newData[i][15] //15-16
      , newData[i][16] //16-17
      , newData[i][4]) //17-18 город

      arrNewData.push(arrNewDataRow)

    }
    else {
      YASHIFTS.getRange(index+1,2).setValue(newData[i][1])
      YASHIFTS.getRange(index+1,3).setValue(newData[i][2])
      YASHIFTS.getRange(index+1,4).setValue(newData[i][3])
      YASHIFTS.getRange(index+1,8).setValue(newData[i][8])
      YASHIFTS.getRange(index+1,9).setValue(newData[i][9])
      YASHIFTS.getRange(index+1,10).setValue(new Date (Date.parse(newData[i][10])).toLocaleDateString('ru-RU'))
      YASHIFTS.getRange(index+1,11).setValue((newData[i][10].split(" "))[1])
      YASHIFTS.getRange(index+1,12).setValue((newData[i][11].split(" "))[1])
      YASHIFTS.getRange(index+1,13).setValue((newData[i][12].split(" "))[1])
      YASHIFTS.getRange(index+1,14).setValue((newData[i][13].split(" "))[1])
      YASHIFTS.getRange(index+1,15).setValue(newData[i][14])
      YASHIFTS.getRange(index+1,16).setValue(newData[i][15])
      YASHIFTS.getRange(index+1,17).setValue(newData[i][16])
      YASHIFTS.getRange(index+1,18).setValue(newData[i][4])
    }
  }

  try {
    let lnght = arrNewData.length
    let arrAA = YASHIFTS.getRange('A:A').getDisplayValues()
    let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
    let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1

    YASHIFTS.insertRowsAfter(YASHIFTS.getMaxRows(), lnght);
    YASHIFTS.getRange(lastRow, 1, lnght, 18).setValues(arrNewData) } catch {console.log('nothing')}
    var time = new Date()
    MONITOR.getRange('C5').setValue("✅ Обновлено " + time)

    REPORT.getRange('A1').activate();
    SpreadsheetApp.enableAllDataSourcesExecution();
    REPORT.refreshAllDataSources();

  Drive.Files.remove(newfileId)
  if (MONITOR.getRange('A1').getValue() > 0) {
    parsing()
  }
}

function insertOperation(newfileId) {
  var NEWFILE = SpreadsheetApp.openById(newfileId).getSheetByName('Sheet1'); //загружаемый файл

  var colNfId = NEWFILE.getRange('C:C').getDisplayValues().filter(clearArray).flat(); //Ид смен из загружаемого файла
  var colYsId = YAOPERATIONS.getRange('C:C').getDisplayValues().filter(clearArray).flat(); //Ид смен из основного файла

  var newData = NEWFILE.getRange('A:Z').getDisplayValues().filter(clearArray); //массив данных из загружаемого файла

  let arrNewData = []

  for(let i = 1; i < colNfId.length; i++) {

    let index = colYsId.indexOf(colNfId[i])

    if (index == -1) {

      let arrNewDataRow = []

      arrNewDataRow.push( 
        new Date (Date.parse(newData[i][0])).toLocaleDateString('ru-RU')
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
      , newData[i][20]
      , newData[i][21]
      , newData[i][22]
      , newData[i][23]
      , newData[i][24]
      , newData[i][25]
      )

      arrNewData.push(arrNewDataRow)

    }
  }

  try {
  let lnght = arrNewData.length
  let arrAA = YAOPERATIONS.getRange('A:A').getDisplayValues()
  let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
  let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1

  YAOPERATIONS.insertRowsAfter(YAOPERATIONS.getMaxRows(), lnght);
  YAOPERATIONS.getRange(lastRow, 1, lnght, 26).setValues(arrNewData) } catch {console.log('nothing')}
  var time = new Date()
  MONITOR.getRange('C5').setValue("✅ Обновлено " + time)

  REPORT.getRange('A1').activate();
  SpreadsheetApp.enableAllDataSourcesExecution();
  REPORT.refreshAllDataSources();

  Drive.Files.remove(newfileId)
  if (MONITOR.getRange('A1').getValue() > 0) {
    parsing()
  }
}

function insertOldReport(newfileId) {
  var NEWFILE = SpreadsheetApp.openById(newfileId).getSheetByName('Sheet1'); //загружаемый файл

  var colNfId = NEWFILE.getRange('A:A').getDisplayValues().filter(clearArray).flat(); //Ид смен из загружаемого файла
  var colYsId = YASHIFTS.getRange('A:A').getDisplayValues().filter(clearArray).flat(); //Ид смен из основного файла

  var newData = NEWFILE.getRange('A:P').getDisplayValues().filter(clearArray); //массив данных из загружаемого файла

  var rows = colYsId.length + 1 //длина основного файла

  for(let i = 1; i < colNfId.length; i++) {

    var index = colYsId.indexOf(colNfId[i])

    if (index == -1) {
      Logger.log("new row" + " " + colNfId[i])
      YASHIFTS.getRange(rows,1).setValue(newData[i][0])
      YASHIFTS.getRange(rows,2).setValue(newData[i][1])
      YASHIFTS.getRange(rows,3).setValue(newData[i][2])
      YASHIFTS.getRange(rows,4).setValue(newData[i][3])
      YASHIFTS.getRange(rows,5).setValue(newData[i][4])
      YASHIFTS.getRange(rows,6).setValue(newData[i][5])
      YASHIFTS.getRange(rows,7).setValue(newData[i][6])
      YASHIFTS.getRange(rows,8).setValue(newData[i][7])
      YASHIFTS.getRange(rows,9).setValue(newData[i][8])
      YASHIFTS.getRange(rows,10).setValue(new Date (Date.parse(newData[i][9])).toLocaleDateString('ru-RU'))
      YASHIFTS.getRange(rows,11).setValue(newData[i][10])
      YASHIFTS.getRange(rows,12).setValue(newData[i][11])
      YASHIFTS.getRange(rows,13).setValue(newData[i][12])
      YASHIFTS.getRange(rows,14).setValue(newData[i][13])
      YASHIFTS.getRange(rows,15).setValue(newData[i][14])
      YASHIFTS.getRange(rows,16).setValue(newData[i][15])
      YASHIFTS.getRange(rows,17).setValue(newData[i][16])
      YASHIFTS.getRange(rows,18).setValue(location(newData[i][4]))
      rows = rows + 1
    }
    else {
      Logger.log(index + " " + colNfId[i])
      YASHIFTS.getRange(index+1,2).setValue(newData[i][1])
      YASHIFTS.getRange(index+1,3).setValue(newData[i][2])
      YASHIFTS.getRange(index+1,4).setValue(newData[i][3])
      YASHIFTS.getRange(index+1,8).setValue(newData[i][7])
      YASHIFTS.getRange(index+1,9).setValue(newData[i][8])
      YASHIFTS.getRange(index+1,10).setValue(new Date (Date.parse(newData[i][9])).toLocaleDateString('ru-RU'))
      YASHIFTS.getRange(index+1,13).setValue(newData[i][12])
      YASHIFTS.getRange(index+1,14).setValue(newData[i][13])
      YASHIFTS.getRange(index+1,15).setValue(newData[i][14])
      YASHIFTS.getRange(index+1,16).setValue(newData[i][15])
      YASHIFTS.getRange(index+1,17).setValue(newData[i][16])
    }
  }
  var time = new Date()
  MONITOR.getRange('C5').setValue("✅ Обновлено " + time)

  REPORT.getRange('A1').activate();
  SpreadsheetApp.enableAllDataSourcesExecution();
  REPORT.refreshAllDataSources();

  Drive.Files.remove(newfileId)
  if (MONITOR.getRange('A1').getValue() > 0) {
    parsing()
  }
}

function location (qst) {
  var ans
  var locList = LOCATIONS.getRange('A2:B200').getValues().filter(clearArray)  
    for (let i = 0; i < locList.length; i++) {
      if (qst.includes(locList[i][0])) {
        ans = locList[i][1]
        break
      }
    }
    for (let i = 0; i < locList.length; i++) {
      if (locList[i][0].includes(qst)) {
        ans = locList[i][1]
        break
      }
    }
  return (ans)
}
