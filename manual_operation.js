const FORM = FormApp.openById('1eAsjg_lHUkR9KsrMcSJoBD4LkEWpRXFg7-_-p4QntZI');
const MANUALOPERATIONS = SpreadsheetApp.openById('1UjYY7p8HbSk1V3iu-mo4mHd3O2_7SqYahVNM4I8PdZ8').getSheetByName('Релокация');
const MONITOR = SpreadsheetApp.openById('1UjYY7p8HbSk1V3iu-mo4mHd3O2_7SqYahVNM4I8PdZ8').getSheetByName('Монитор');

function clearArray(ar) {
  return (ar != null && ar != "" || ar === 0);
}

function parsing() {
  var t = MONITOR.getRange('A1').getValue()
  if(MONITOR.getRange('B6').getValue() == "⌛ Идет загрузка") {
    return
  }
  else {
    t > 0 ? MONITOR.getRange('A1').setValue(t-1) :
    MONITOR.insertRowBefore(6)
    MONITOR.getRange('B6').setValue("⌛ Идет загрузка")
    let responses = FORM.getResponses();
    let fileId = responses[responses.length-1-t].getItemResponses()[0].getResponse()
    var file = DriveApp.getFileById(fileId);
    let blob = file.getBlob().setContentType(MimeType.MICROSOFT_EXCEL)
      let resource = {
      title: blob.getName().split('.')[0],
      mimeType: MimeType.GOOGLE_SHEETS
    };
    let newfileId = Drive.Files.insert(resource, blob, {"convert":"true"}).getId();
    DriveApp.getFileById(newfileId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT)

    let fileTitle = JSON.stringify(SpreadsheetApp.openById(newfileId).getSheetByName('Sheet1').getRange('A1:Z1').getValues()[0])
    let yandexOperations = JSON.stringify(["Дата","ID исполнителя","ID слота","ФИО исполнителя","Регион","Транспортное средство","Специализация","Вендор","Продолжительность смены","Своя машина","Завершённые миссии по замене аккумуляторов","Запланированные работы по замене аккумуляторов","Завершённые работы по замене аккумуляторов","Пропущенные работы по замене аккумуляторов","Завершённые миссии по отвозу на склад","Запланированные работы по отвозу на склад","Завершённые работы по отвозу на склад","Пропущенные работы по отвозу на склад","Завершённые миссии по релокации","Запланированные работы по релокации","Завершённые работы по релокации","Пропущенные работы по релокации","Завершённые миссии по вывозу со склада","Запланированные работы по вывозу со склада","Завершённые работы по вывозу со склада","Пропущенные работы по вывозу со склада"])

    if (fileTitle == yandexOperations) {
      MONITOR.getRange('E6').setValue(file.getUrl())
      updateFile(newfileId)
    }
    else {
      MONITOR.getRange('E6').setValue(file.getUrl())
      MONITOR.getRange('B6').setValue("❌ Загружен неверный файл (" + file.getName() + ")")
    }
  }
}

function minMaxDate (dateArray) {
  dateArray = dateArray.filter(clearArray)
  if (dateArray.filter(clearArray)[1][0].includes('.') == true) {
    let day = parseInt(dateArray[1][0].split(' - ')[0].split('.')[0], 10);
    let month = parseInt(dateArray[1][0].split(' - ')[0].split('.')[1], 10) - 1;
    let year = parseInt(dateArray[1][0].split(' - ')[0].split('.')[2], 10);
    var minDate = new Date(year, month, day);
    var maxDate = new Date(year, month, day);


    for (let d = 1; d < dateArray.length; d++) {
      try {

      let day1 = parseInt(dateArray[d][0].split(' - ')[0].split('.')[0], 10);
      let month1 = parseInt(dateArray[d][0].split(' - ')[0].split('.')[1], 10) - 1;
      let year1 = parseInt(dateArray[d][0].split(' - ')[0].split('.')[2], 10);
      let currentDate1 = new Date(year1, month1, day1);

      if(currentDate1 <= minDate) {
        minDate = currentDate1
      }

      let day2 = parseInt(dateArray[d][0].split(' - ')[1].split('.')[0], 10);
      let month2 = parseInt(dateArray[d][0].split(' - ')[1].split('.')[1], 10) - 1;
      let year2 = parseInt(dateArray[d][0].split(' - ')[1].split('.')[2], 10);
      let currentDate2 = new Date(year2, month2, day2);

      if(currentDate2 >= maxDate) {
        maxDate = currentDate2
      }
      }
      catch { Logger.log("ошибочка")}
    }
  }
  else if (dateArray[1].includes('.') == false) {

    let day = parseInt(dateArray[1][0].split('-')[2], 10);
    let month = parseInt(dateArray[1][0].split('-')[1], 10) - 1;
    let year = parseInt(dateArray[1][0].split('-')[0], 10);
    var minDate = new Date(year, month, day);
    var maxDate = new Date(year, month, day);

    for (let d = 1; d < dateArray.length; d++) {
      try {

      let day1 = parseInt(dateArray[d][0].split('-')[2], 10);
      let month1 = parseInt(dateArray[d][0].split('-')[1], 10) - 1;
      let year1 = parseInt(dateArray[d][0].split('-')[0], 10);
      let currentDate1 = new Date(year1, month1, day1);

      if(currentDate1 <= minDate) {
        minDate = currentDate1
      }

      let day2 = parseInt(dateArray[d][0].split('-')[2], 10);
      let month2 = parseInt(dateArray[d][0].split('-')[1], 10) - 1;
      let year2 = parseInt(dateArray[d][0].split('-')[0], 10);
      let currentDate2 = new Date(year2, month2, day2);

      if(currentDate2 >= maxDate) {
        maxDate = currentDate2
      }
      }
      catch { Logger.log("ошибочка")}
    }
  }

  return ([minDate.toLocaleDateString('ru-RU'), maxDate.toLocaleDateString('ru-RU')])
}

function updateFile (newfileId) {
  let NEWFILE = SpreadsheetApp.openById(newfileId).getSheetByName('Sheet1'); //загружаемый файл
  let minDateArray = NEWFILE.getRange('A:A').getDisplayValues();
  let minDate = minMaxDate(minDateArray)[0]
  MONITOR.getRange('C6').setValue(minDate)
  let indexStart = MANUALOPERATIONS.getRange('C:C').getDisplayValues().flat().indexOf(minDate)
  try {
    let arrAA = MANUALOPERATIONS.getRange('C:C').getDisplayValues()
    let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
    var lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1 - indexStart
  } catch {console.log('nothing')}
  let manualOperationsId = MANUALOPERATIONS.getRange(indexStart+1,1,lastRow,1).getDisplayValues()
  let manualOperationsDoers = MANUALOPERATIONS.getRange(indexStart+1,2,lastRow,1).getDisplayValues()
  let manualOperationsDate = MANUALOPERATIONS.getRange(indexStart+1,3,lastRow,1).getDisplayValues()
  let manualOperationsCity = MANUALOPERATIONS.getRange(indexStart+1,5,lastRow,1).getDisplayValues()
  let manualOperationsType = MANUALOPERATIONS.getRange(indexStart+1,6,lastRow,1).getDisplayValues()
  let manualOperationsCount = MANUALOPERATIONS.getRange(indexStart+1,7,lastRow,1).getDisplayValues()
  let file = NEWFILE.getRange("A:Z").getDisplayValues()
  for (let f = 0; f < manualOperationsDate.length; f++) {
    for (let d = 0; d < file.length; d++) {

      let firstdate = file[d][0]
      let day = parseInt(firstdate.split('-')[2], 10);
      let month = parseInt(firstdate.split('-')[1], 10) - 1;
      let year = parseInt(firstdate.split('-')[0], 10);
      let currentDate = new Date(year, month, day);
      let date = currentDate.toLocaleDateString('ru-RU')

      if(file[d][1].includes(manualOperationsId[f][0])
      && date.includes(manualOperationsDate[f][0])) {
        if (manualOperationsType[f][0] == 'Сбор') {
          let pick_up = parseInt(NEWFILE.getRange(d+1, 28, 1, 1).getValue()) > 0 ? parseInt(manualOperationsCount[f][0]) + parseInt(NEWFILE.getRange(d+1, 28, 1, 1).getValue()) : parseInt(manualOperationsCount[f][0])
          NEWFILE.getRange(d+1, 28, 1, 1).setValue(pick_up)
        }
        if (manualOperationsType[f][0] == 'Перезарядка') {
          let recharge = parseInt(NEWFILE.getRange(d+1, 29, 1, 1).getValue()) > 0 ? parseInt(manualOperationsCount[f][0]) + parseInt(NEWFILE.getRange(d+1, 29, 1, 1).getValue()) : parseInt(manualOperationsCount[f][0])
          NEWFILE.getRange(d+1, 29, 1, 1).setValue(recharge)
        }
        if (manualOperationsType[f][0] == 'Вывоз') {
          let deploy = parseInt(NEWFILE.getRange(d+1, 30, 1, 1).getValue()) > 0 ? parseInt(manualOperationsCount[f][0]) + parseInt(NEWFILE.getRange(d+1, 30, 1, 1).getValue()) : parseInt(manualOperationsCount[f][0])
          NEWFILE.getRange(d+1, 30, 1, 1).setValue(deploy)
        }
        if (manualOperationsType[f][0] == 'Релокация') {
          let relocations = parseInt(NEWFILE.getRange(d+1, 31, 1, 1).getValue()) > 0 ? parseInt(manualOperationsCount[f][0]) + parseInt(NEWFILE.getRange(d+1, 31, 1, 1).getValue()) : parseInt(manualOperationsCount[f][0])
          NEWFILE.getRange(d+1, 31, 1, 1).setValue(relocations)
        }
        done = 1
        break
      }
      else if (file[d][3].includes(manualOperationsDoers[f][0]) 
        && file[d][4].includes(manualOperationsCity[f][0]) 
        && date.includes(manualOperationsDate[f][0]) 
        && manualOperationsDoers[f][0] != '') {
        if (manualOperationsType[f][0] == 'Сбор') {
          let pick_up = parseInt(NEWFILE.getRange(d+1, 28, 1, 1).getValue()) > 0 ? parseInt(manualOperationsCount[f][0]) + parseInt(NEWFILE.getRange(d+1, 28, 1, 1).getValue()) : parseInt(manualOperationsCount[f][0])
          NEWFILE.getRange(d+1, 28, 1, 1).setValue(pick_up)
        }
        if (manualOperationsType[f][0] == 'Перезарядка') {
          let recharge = parseInt(NEWFILE.getRange(d+1, 29, 1, 1).getValue()) > 0 ? parseInt(manualOperationsCount[f][0]) + parseInt(NEWFILE.getRange(d+1, 29, 1, 1).getValue()) : parseInt(manualOperationsCount[f][0])
          NEWFILE.getRange(d+1, 29, 1, 1).setValue(recharge)
        }
        if (manualOperationsType[f][0] == 'Вывоз') {
          let deploy = parseInt(NEWFILE.getRange(d+1, 30, 1, 1).getValue()) > 0 ? parseInt(manualOperationsCount[f][0]) + parseInt(NEWFILE.getRange(d+1, 30, 1, 1).getValue()) : parseInt(manualOperationsCount[f][0])
          NEWFILE.getRange(d+1, 30, 1, 1).setValue(deploy)
        }
        if (manualOperationsType[f][0] == 'Релокация') {
          let relocations = parseInt(NEWFILE.getRange(d+1, 31, 1, 1).getValue()) > 0 ? parseInt(manualOperationsCount[f][0]) + parseInt(NEWFILE.getRange(d+1, 31, 1, 1).getValue()) : parseInt(manualOperationsCount[f][0])
          NEWFILE.getRange(d+1, 31, 1, 1).setValue(relocations)
        }
        done = 1
        break
      }
    }
  }
  let newURl = DriveApp.getFileById(newfileId).getUrl()
  MONITOR.getRange('F6').setValue(newURl)
  MONITOR.getRange('B6').setValue("✅ Готово")
  NEWFILE.getRange(1,27,1,1).setValue("Общая продолжительность ручной смены")
  NEWFILE.getRange(1,28,1,1).setValue("Завершённые работы по ручному отвозу на склад")
  NEWFILE.getRange(1,29,1,1).setValue("Завершённые работы по ручной замене аккумуляторов")
  NEWFILE.getRange(1,30,1,1).setValue("Завершённые работы по ручному вывозу")
  NEWFILE.getRange(1,31,1,1).setValue("Завершённые работы по ручной релокации")
}

