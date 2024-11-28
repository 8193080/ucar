const TEST = SpreadsheetApp.openById('1mNB5OzeHMKv1vtSy6YB9dPev8qU6FjEm8sI7jSGOjAk').getSheetByName('test');

function demandUrent() {

  let date = NEWDEMAND.getRange('AA26:AG26').getDisplayValues().flat()
  let shiftType = ['day', 'night']
  let roles = ['Скаут', 'Водитель', 'Мастер приемщик']
  let cities = ['Томск', 'Новосибирск', 'Нижний Новгород', 'Казань', 'Самара']

  let scoutsTomsk = NEWDEMAND.getRange('AA28:AG29').getValues()
  let driversTomsk = NEWDEMAND.getRange('AA30:AG31').getValues()
  let mastersTomsk = NEWDEMAND.getRange('AA32:AG33').getValues()

  let scoutsNovosibirsk = NEWDEMAND.getRange('AA35:AG36').getValues()
  let driversNovosibirsk = NEWDEMAND.getRange('AA37:AG38').getValues()
  let mastersNovosibirsk = NEWDEMAND.getRange('AA39:AG40').getValues()

  let scoutsGorkiy = NEWDEMAND.getRange('AA42:AG43').getValues()
  let driversGorkiy = NEWDEMAND.getRange('AA44:AG45').getValues()
  let mastersGorkiy = NEWDEMAND.getRange('AA46:AG47').getValues()

  let scoutsKazan = NEWDEMAND.getRange('AA49:AG50').getValues()
  let driversKazan = NEWDEMAND.getRange('AA51:AG52').getValues()
  let mastersKazan = NEWDEMAND.getRange('AA53:AG54').getValues()

  let scoutsSamara = NEWDEMAND.getRange('AA56:AG57').getValues()
  let driversSamara = NEWDEMAND.getRange('AA58:AG59').getValues()
  let mastersSamara = NEWDEMAND.getRange('AA60:AG61').getValues()

  let data = []
  let finalData = []
  let t = 0

  data.push(
    scoutsTomsk,
    driversTomsk,
    mastersTomsk,
    scoutsNovosibirsk,
    driversNovosibirsk,
    mastersNovosibirsk,
    scoutsGorkiy,
    driversGorkiy,
    mastersGorkiy,
    scoutsKazan,
    driversKazan,
    mastersKazan,
    scoutsSamara,
    driversSamara,
    mastersSamara
    )

  for (let c = 0; c < cities.length; c++) {
    for (let r = 0; r < roles.length; r++) {    
      for (let s = 0; s < shiftType.length; s++) {      
          for (let d = 0; d < date.length; d++) {
            let temporaryData = []        
            temporaryData.push(date[d], shiftType[s], cities[c], roles[r], data[t][s][d],  date[d] + shiftType[s] + cities[c] + roles[r])
            finalData.push(temporaryData)
          }
        }
      t++
    }
  }

  let newData = []
  let archive = NEWARCHIVE.getRange('AH:AH').getDisplayValues().flat()
  for (let d = 0; d < finalData.length; d++) {
    let index = -1 
    index = archive.indexOf(finalData[d][5])
    let arr = []
    if (index != -1 && finalData[d][4] != 0) {
      arr.push(finalData[d])
      NEWARCHIVE.getRange(index+1,29,1,6).setValues(arr)
    }
    else if (index == -1 && finalData[d][4] != 0){
      arr.push(finalData[d])
      newData.push(finalData[d])
    }
    
  }
  if (newData.length > 0) {
    //try {
      let lnght = newData.length
      let arrAA = NEWARCHIVE.getRange('AC:AC').getDisplayValues()
      let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
      let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1
      NEWARCHIVE.insertRowsAfter(NEWARCHIVE.getMaxRows(), lnght);
      NEWARCHIVE.getRange(lastRow, 29, lnght, 6).setValues(newData)
    //} catch {}
  }
  SpreadsheetApp.getUi().alert('Данные внесены'); 
  NEWDEMAND.getRange('AA28:AG61').clear({contentsOnly: true}); //очищаем только значения в ячейках
}

function demandVeloBike() {

  let date = NEWDEMAND.getRange('AA2:AG2').getDisplayValues().flat()
  let shiftType = ['day', 'night']
  let roles = ['Вело Чарджер Велобайк', 'Водитель Велобайк', 'Чарджер Велобайк', 'Скаут']

  let chargersAll = NEWDEMAND.getRange('AA4:AG5').getValues()
  let driversAll = NEWDEMAND.getRange('AA6:AG7').getValues()
  let carchargersAll = NEWDEMAND.getRange('AA8:AG9').getValues()
  let scoutsAll = NEWDEMAND.getRange('AA10:AG11').getValues()

  let chargersInside = NEWDEMAND.getRange('AA13:AG14').getValues()
  let driversInside = NEWDEMAND.getRange('AA15:AG16').getValues()
  let carchargersInside = NEWDEMAND.getRange('AA17:AG18').getValues()
  let scoutsInside = NEWDEMAND.getRange('AA19:AG20').getValues()

  let data = []
  let finalData = []
  let t = 0

  data.push(chargersAll, chargersInside, driversAll, driversInside, carchargersAll, carchargersInside, scoutsAll, scoutsInside)

  Logger.log(data)

  for (let r = 0; r < roles.length; r++) {
    for (let s = 0; s < shiftType.length; s++) {
      for (let d = 0; d < date.length; d++) {
        let temporaryData = []        
        temporaryData.push(date[d], shiftType[s], roles[r], data[t][s][d], data[t+1][s][d],  date[d] + shiftType[s] + roles[r])
        finalData.push(temporaryData)
      }
    }
    t=t+2
  }
  Logger.log(finalData)

  let newData = []
  let archive = NEWARCHIVE.getRange('AA:AA').getDisplayValues().flat()
  for (let d = 0; d < finalData.length; d++) {
    let index = -1 
    index = archive.indexOf(finalData[d][5])
    let arr = []
    if (index != -1 && (finalData[d][4] != 0 || finalData[d][3] != 0)) {
      arr.push(finalData[d])
      NEWARCHIVE.getRange(index+1,22,1,6).setValues(arr)
    }
    else if (index == -1 && (finalData[d][4] != 0 || finalData[d][3] != 0)){
      arr.push(finalData[d])
      newData.push(finalData[d])
    }
    
  }
  if (newData.length > 0) {
    try {
      let lnght = newData.length
      let arrAA = NEWARCHIVE.getRange('V:V').getDisplayValues()
      let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
      let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1
      NEWARCHIVE.insertRowsAfter(NEWARCHIVE.getMaxRows(), lnght);
      NEWARCHIVE.getRange(lastRow, 22, lnght, 6).setValues(newData) 
    } catch {}
  }
  SpreadsheetApp.getUi().alert('Данные внесены');
  NEWDEMAND.getRange('AA4:AG12').clear({contentsOnly: true}); //очищаем только значения в ячейках
}


