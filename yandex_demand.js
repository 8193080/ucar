const NEED = SpreadsheetApp.openById('1mNB5OzeHMKv1vtSy6YB9dPev8qU6FjEm8sI7jSGOjAk').getSheetByName('потребность');
const NEWARCHIVE = SpreadsheetApp.openById('1mNB5OzeHMKv1vtSy6YB9dPev8qU6FjEm8sI7jSGOjAk').getSheetByName('archive_demand');
const NEWDEMAND = SpreadsheetApp.openById('1mNB5OzeHMKv1vtSy6YB9dPev8qU6FjEm8sI7jSGOjAk').getSheetByName('внесение потребности');
const currentNeed = NEED.getRange("A:E").getValues().filter(clearArray);
const arrayCities = ['Москва','Санкт-Петербург','Тула','Краснодар','Екатеринбург','Казань']

function clearArray(ar) {
  return (ar != null && ar != "" || ar === 0);
}

var arrayCurrentNeed = []
var arrayArchiveNeed = []

function checkNeed() { //проверка внесенали уже потребность
  var cNlength = currentNeed.length - 1
  for (let c = 0; c < arrayCities.length; c++) {
    var maxDate = currentNeed[cNlength][0]
    for (let i = cNlength; i >= 0; i--) {
      if (currentNeed[i][1] == arrayCities[c]) {
        maxDate = maxDate < currentNeed[i][0] ? currentNeed[i][0] : maxDate
      }
    }

    var arr = [maxDate, arrayCities[c]]
    arrayCurrentNeed.push(arr)
  }
  return (arrayCurrentNeed)
}

function checkArchiveNeed() {
  var cNlength = archiveNeed.length - 1
  for (let c = 0; c < arrayCities.length; c++) {
    var maxDate = archiveNeed[cNlength][0]
    for (let i = cNlength; i >= 0; i--) {
      if (archiveNeed[i][1] == arrayCities[c]) {
        maxDate = maxDate < archiveNeed[i][0] ? archiveNeed[i][0] : maxDate
      }
    }
    var arr = [maxDate, arrayCities[c]]
    arrayArchiveNeed.push(arr)
  }
  return (arrayArchiveNeed)
}

function refresh() {
  var cn = checkNeed()
  var can = checkArchiveNeed()
  Logger.log(cn)
  Logger.log(can)
  for (let i = 0; i < cn.length; i++) {
    var city = cn[i][0] > can[i][0] ? cn[i][1] : 'empty'
    if (city != 'empty') {
      Logger.log(city)
      Logger.log(can[i][0])
      for (a = 0; a < currentNeed.length; a++) {
        if (currentNeed[a][1] == city && currentNeed[a][0] > can[i][0]) {
          ARCHIVE.insertRowsBefore(2,1);
          //ARCHIVE.getRange('A2:E2').setValues(currentNeed[a])
          ARCHIVE.getRange('A2').setValue(currentNeed[a][0])
          ARCHIVE.getRange('B2').setValue(currentNeed[a][1])
          ARCHIVE.getRange('C2').setValue(currentNeed[a][2])
          ARCHIVE.getRange('D2').setValue(currentNeed[a][3])
          ARCHIVE.getRange('E2').setValue(currentNeed[a][4])
        }
      }
    }
  }
}

function newDemand () { //внесение потребности ежендевно - см. триггер на скрипт
  
  let scooterDay = NEWDEMAND.getRange('C3:L12').getValues()
  let scooterNight = NEWDEMAND.getRange('C15:L24').getValues() //энды

  let driversDay = NEWDEMAND.getRange('C28:L37').getValues()
  let driversNight = NEWDEMAND.getRange('C40:L49').getValues() //водители

  let whDay = NEWDEMAND.getRange('C55:L64').getValues()
  let whNight = NEWDEMAND.getRange('C66:L75').getValues() //мастера смены

  let whAddDay = NEWDEMAND.getRange('C77:L86').getValues()
  let whAddNight = NEWDEMAND.getRange('C88:L97').getValues() //муверы
  
  let repairDay = NEWDEMAND.getRange('C99:L108').getValues()
  let repairNight = NEWDEMAND.getRange('C110:L119').getValues() //складские ремонтники
  
  let carChargerDay = NEWDEMAND.getRange('C124:L133').getValues()
  let carChargerNight = NEWDEMAND.getRange('C135:L144').getValues() //энды на авто

  let carRelocatorDay = NEWDEMAND.getRange('C146:L155').getValues()
  let carRelocatorNight = NEWDEMAND.getRange('C157:L166').getValues() //релокаторы

  let carRebalancerDay = NEWDEMAND.getRange('C168:L177').getValues()
  let carRebalancerNight = NEWDEMAND.getRange('C179:L188').getValues() //ребалансировщики

  let finalData = []

  let date = NEWDEMAND.getRange('C2:L2').getDisplayValues().flat()
  let cities = NEWDEMAND.getRange('B3:B12').getDisplayValues().flat()
  let roles = ['scooter', 'car', 'warehouse', 'add_wh', 'repair', 'car_charger', 'car_relocator', 'car_rebalancer']
  let shiftType = ['day', 'night']

  let data = []
  data.push(
      scooterDay, scooterNight
    , driversDay, driversNight
    , whDay, whNight
    , whAddDay, whAddNight
    , repairDay, repairNight
    , carChargerDay, carChargerNight
    , carRelocatorDay, carRelocatorNight
    , carRebalancerDay, carRebalancerNight
    )
  let t = 0

  for (let r = 0; r < roles.length; r++) {
    for (let s = 0; s < shiftType.length; s++) {
      for (let d = 0; d < date.length; d++) {
        for (let c = 0; c < cities.length; c++) {
          let temporaryData = []
          temporaryData.push(date[d], shiftType[s], cities[c], roles[r], data[t][c][d],  date[d] + shiftType[s] + cities[c] + roles[r])
          finalData.push(temporaryData)
        }
      }
      t++
    }
  }
  let newData = []
  let archive = NEWARCHIVE.getRange('F:F').getDisplayValues().flat()
  for (let d = 0; d < finalData.length; d++) {
    let index = -1 
    index = archive.indexOf(finalData[d][5])
    let arr = []
    if (index != -1 && finalData[d][4] != 0) {
      arr.push(finalData[d])
      NEWARCHIVE.getRange(index+1,1,1,6).setValues(arr)
    }
    else if (index == -1 && finalData[d][4] != 0){
      arr.push(finalData[d])
      newData.push(finalData[d])
    }
    
  }
  if (newData.length > 0) {
    try {
      let lnght = newData.length
      let arrAA = NEWARCHIVE.getRange('A:A').getDisplayValues()
      let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
      let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1
      NEWARCHIVE.insertRowsAfter(NEWARCHIVE.getMaxRows(), lnght);
      NEWARCHIVE.getRange(lastRow, 1, lnght, 6).setValues(newData) 
    } catch {}
  }
}

function newDemandThu () { //внесение опубликованных смен в четверг - см. триггер на скрипт
  
  let scooterDay = NEWDEMAND.getRange('C3:L12').getValues()
  let scooterNight = NEWDEMAND.getRange('C15:L24').getValues() //энды

  let driversDay = NEWDEMAND.getRange('C28:L37').getValues()
  let driversNight = NEWDEMAND.getRange('C40:L49').getValues() //водители

  let whDay = NEWDEMAND.getRange('C55:L64').getValues()
  let whNight = NEWDEMAND.getRange('C66:L75').getValues() //мастера смены

  let whAddDay = NEWDEMAND.getRange('C77:L86').getValues()
  let whAddNight = NEWDEMAND.getRange('C88:L97').getValues() //муверы
  
  let repairDay = NEWDEMAND.getRange('C99:L108').getValues()
  let repairNight = NEWDEMAND.getRange('C110:L119').getValues() //складские ремонтники
  
  let carChargerDay = NEWDEMAND.getRange('C124:L133').getValues()
  let carChargerNight = NEWDEMAND.getRange('C135:L144').getValues() //энды на авто

  let carRelocatorDay = NEWDEMAND.getRange('C146:L155').getValues()
  let carRelocatorNight = NEWDEMAND.getRange('C157:L166').getValues() //релокаторы

  let carRebalancerDay = NEWDEMAND.getRange('C168:L177').getValues()
  let carRebalancerNight = NEWDEMAND.getRange('C179:L188').getValues() //ребалансировщики

  let finalData = []

  let date = NEWDEMAND.getRange('C2:L2').getDisplayValues().flat()
  let cities = NEWDEMAND.getRange('B3:B12').getDisplayValues().flat()
  let roles = ['scooter', 'car', 'warehouse', 'add_wh', 'repair', 'car_charger', 'car_relocator', 'car_rebalancer']
  let shiftType = ['day', 'night']

  let data = []
    data.push(
      scooterDay, scooterNight
    , driversDay, driversNight
    , whDay, whNight
    , whAddDay, whAddNight
    , repairDay, repairNight
    , carChargerDay, carChargerNight
    , carRelocatorDay, carRelocatorNight
    , carRebalancerDay, carRebalancerNight
    )
  let t = 0

  for (let r = 0; r < roles.length; r++) {
    for (let s = 0; s < shiftType.length; s++) {
      for (let d = 0; d < date.length; d++) {
        for (let c = 0; c < cities.length; c++) {
          let temporaryData = []
          temporaryData.push(date[d], shiftType[s], cities[c], roles[r], data[t][c][d],  date[d] + shiftType[s] + cities[c] + roles[r])
          finalData.push(temporaryData)
        }
      }
      t++
    }
  }
  let newData = []
  let archive = NEWARCHIVE.getRange('M:M').getDisplayValues().flat()
  for (let d = 0; d < finalData.length; d++) {
    let index = -1 
    index = archive.indexOf(finalData[d][5])
    let arr = []
    if (index == -1 && finalData[d][4] != 0){
      arr.push(finalData[d])
      newData.push(finalData[d])
      //row++
    }
    
  }
  if (newData.length > 0) {
    try {
      let lnght = newData.length
      let arrAA = NEWARCHIVE.getRange('H:H').getDisplayValues()
      let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
      let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1
      NEWARCHIVE.insertRowsAfter(NEWARCHIVE.getMaxRows(), lnght);
      NEWARCHIVE.getRange(lastRow, 8, lnght, 6).setValues(newData) 
  } catch {}
  }
  //NEWARCHIVE.getRange(2,1,finalData.length,6).setValues(finalData)
}

function newSignUp () { //внесение записей
  
  let scooterDay = NEWDEMAND.getRange('P3:X12').getValues()
  let scooterNight = NEWDEMAND.getRange('P15:X24').getValues() //энды

  let driversDay = NEWDEMAND.getRange('P28:X37').getValues()
  let driversNight = NEWDEMAND.getRange('P40:X49').getValues() //водители

  let whDay = NEWDEMAND.getRange('P55:X64').getValues()
  let whNight = NEWDEMAND.getRange('P66:X75').getValues() //мастера смены

  let whAddDay = NEWDEMAND.getRange('P77:X86').getValues()
  let whAddNight = NEWDEMAND.getRange('P88:X97').getValues() //муверы
  
  let repairDay = NEWDEMAND.getRange('P99:X108').getValues()
  let repairNight = NEWDEMAND.getRange('P110:X119').getValues() //складские ремонтники
  
  let carChargerDay = NEWDEMAND.getRange('P124:X133').getValues()
  let carChargerNight = NEWDEMAND.getRange('P135:X144').getValues() //энды на авто

  let carRelocatorDay = NEWDEMAND.getRange('P146:X155').getValues()
  let carRelocatorNight = NEWDEMAND.getRange('P157:X166').getValues() //релокаторы

  let carRebalancerDay = NEWDEMAND.getRange('P168:X177').getValues()
  let carRebalancerNight = NEWDEMAND.getRange('P179:X188').getValues() //ребалансировщики


  let finalData = []

  let date = NEWDEMAND.getRange('P2:X2').getDisplayValues().flat()
  let cities = NEWDEMAND.getRange('B3:B12').getDisplayValues().flat()
  let roles = ['scooter', 'car', 'warehouse', 'add_wh', 'repair', 'car_charger', 'car_relocator', 'car_rebalancer']
  let shiftType = ['day', 'night']

  let data = []
    data.push(
      scooterDay, scooterNight
    , driversDay, driversNight
    , whDay, whNight
    , whAddDay, whAddNight
    , repairDay, repairNight
    , carChargerDay, carChargerNight
    , carRelocatorDay, carRelocatorNight
    , carRebalancerDay, carRebalancerNight
    )
  let t = 0

  for (let r = 0; r < roles.length; r++) {
    for (let s = 0; s < shiftType.length; s++) {
      for (let d = 0; d < date.length; d++) {
        for (let c = 0; c < cities.length; c++) {
          let temporaryData = []
          temporaryData.push(date[d], shiftType[s], cities[c], roles[r], data[t][c][d],  date[d] + shiftType[s] + cities[c] + roles[r])
          finalData.push(temporaryData)
        }
      }
      t++
    }
  }
  let newData = []
  let archive = NEWARCHIVE.getRange('T:T').getDisplayValues().flat()
  for (let d = 0; d < finalData.length; d++) {
    let index = -1 
    index = archive.indexOf(finalData[d][5])
    let arr = []
    if (index != -1 && finalData[d][4] != 0) {
      arr.push(finalData[d])
      NEWARCHIVE.getRange(index+1,15,1,6).setValues(arr)
    }
    else if (index == -1 && finalData[d][4] != 0){
      arr.push(finalData[d])
      newData.push(finalData[d])
    }
    
  }
  if (newData.length > 0) {
    try {
      let lnght = newData.length
      let arrAA = NEWARCHIVE.getRange('O:O').getDisplayValues()
      let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
      let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1
      NEWARCHIVE.insertRowsAfter(NEWARCHIVE.getMaxRows(), lnght);
      NEWARCHIVE.getRange(lastRow, 15, lnght, 6).setValues(newData) 
    } catch {}
  }
}

function newManualDemand () { //внесение ручной потребности

  let scooterDay = NEWDEMAND.getRange('L3:L12').getValues()
  let scooterNight = NEWDEMAND.getRange('L15:L24').getValues() //энды

  let driversDay = NEWDEMAND.getRange('L28:L37').getValues()
  let driversNight = NEWDEMAND.getRange('L40:L49').getValues() //водители

  let whDay = NEWDEMAND.getRange('L55:L64').getValues()
  let whNight = NEWDEMAND.getRange('L66:L75').getValues() //мастера смены

  let whAddDay = NEWDEMAND.getRange('L77:L86').getValues()
  let whAddNight = NEWDEMAND.getRange('L88:L97').getValues() //муверы
  
  let repairDay = NEWDEMAND.getRange('L99:L108').getValues()
  let repairNight = NEWDEMAND.getRange('L110:L119').getValues() //складские ремонтники
  
  let carChargerDay = NEWDEMAND.getRange('L124:L133').getValues()
  let carChargerNight = NEWDEMAND.getRange('L135:L144').getValues() //энды на авто

  let carRelocatorDay = NEWDEMAND.getRange('L146:L155').getValues()
  let carRelocatorNight = NEWDEMAND.getRange('L157:L166').getValues() //релокаторы

  let carRebalancerDay = NEWDEMAND.getRange('L168:L177').getValues()
  let carRebalancerNight = NEWDEMAND.getRange('L179:L188').getValues() //ребалансировщики


  let finalData = []

  let date = NEWDEMAND.getRange('L2:L2').getDisplayValues().flat()
  let cities = NEWDEMAND.getRange('B3:B12').getDisplayValues().flat()
  let roles = ['scooter', 'car', 'warehouse', 'add_wh', 'repair', 'car_charger', 'car_relocator', 'car_rebalancer']
  let shiftType = ['day', 'night']

  let data = []
  data.push(
      scooterDay, scooterNight
    , driversDay, driversNight
    , whDay, whNight
    , whAddDay, whAddNight
    , repairDay, repairNight
    , carChargerDay, carChargerNight
    , carRelocatorDay, carRelocatorNight
    , carRebalancerDay, carRebalancerNight
  )
  let t = 0

  let info = 'Нет новых данных'

  for (let r = 0; r < roles.length; r++) {
    for (let s = 0; s < shiftType.length; s++) {
      for (let d = 0; d < date.length; d++) {
        for (let c = 0; c < cities.length; c++) {
          let temporaryData = []
          temporaryData.push(date[d], shiftType[s], cities[c], roles[r], data[t][c][d],  date[d] + shiftType[s] + cities[c] + roles[r])
          finalData.push(temporaryData)
        }
      }
      t++
    }
  }
  let row = 1
  let newData = []
  let archive = NEWARCHIVE.getRange('F:F').getDisplayValues().flat()
  for (let d = 0; d < finalData.length; d++) {
    let index = archive.indexOf(finalData[d][5])
    let arr = []
    if (index != -1 && finalData[d][4] != 0) {
      arr.push(finalData[d])
      NEWARCHIVE.getRange(index+row,1,1,6).setValues(arr)
    }
    else if (index == -1 && finalData[d][4] != 0){
      arr.push(finalData[d])
      newData.push(finalData[d])
    }
  }
  if (newData.length > 0) {
    //try {
      let lnght = newData.length
      let arrAA = NEWARCHIVE.getRange('F:F').getDisplayValues()
      let countSpace = arrAA.reverse().findIndex(row => row.join('') !== '');
      let lastRow = arrAA.length - (countSpace === -1 ? arrAA.length : countSpace) + 1
      NEWARCHIVE.insertRowsAfter(NEWARCHIVE.getMaxRows(), lnght);
      NEWARCHIVE.getRange(lastRow, 1, lnght, 6).setValues(newData)
    //} catch {}
  }
  SpreadsheetApp.getUi().alert('Данные внесены');
  NEWDEMAND.getRange('L3:L188').clear({contentsOnly: true});
}
