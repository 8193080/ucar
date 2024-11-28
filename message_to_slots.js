const DATA = SpreadsheetApp.openById('12Vs6IdBk2Pq_rCcebrIHRuzsCyLpo9XxYz7p9nJWIz4').getSheetByName('data')
const MESSAGE = SpreadsheetApp.openById('12Vs6IdBk2Pq_rCcebrIHRuzsCyLpo9XxYz7p9nJWIz4').getSheetByName('message')

function clearArray(ar) {
  return (ar != null && ar != "" || ar === 0);
}

function message() {
  var dateArray = DATA.getRange('A:A').getValues().filter(clearArray)
  var startDate = MESSAGE.getRange('C2').getValue()
  var finishDate = MESSAGE.getRange('D2').getValue()
  var distinctLoc = []
  var distinctDate = []
  var chkDate = 0
  var message

  var array = DATA.getRange('A:F').getValues().filter(clearArray)

  if (!(startDate >= finishDate)) {
    Logger.log('true')
    message = 'Распределение с ' + startDate.toLocaleDateString('ru-RU',{ day: '2-digit', month: '2-digit' }) + ' по ' + finishDate.toLocaleDateString('ru-RU',{ day: '2-digit', month: '2-digit' })
  }
  else {
    Logger.log('false')
    message = 'Распределение на ' + startDate.toLocaleDateString('ru-RU',{ day: '2-digit', month: '2-digit' })
  }

  var location 
  var date

  for (let i = dateArray.length -1; i >= 0 ; i--) {
    if (dateArray[i][0] >= startDate && dateArray[i][0] <= finishDate) {
      if (!distinctLoc.includes(array[i][1])) {
        distinctLoc.push(array[i][1])
      }
    }
    if (dateArray[i][0] >= startDate && dateArray[i][0] <= finishDate) {
      if (!distinctDate.includes(array[i][0])) {
        distinctDate.push(array[i][0])
      }
    }
  }
  Logger.log(dateArray.length)
  for (let i = 0; i < distinctLoc.length; i++) {

    message = message + "\n\n📍" + distinctLoc[i]

    for (let d = 0; d < distinctDate.length; d++){

      chkDate = 0

      for (let l = 0; l < dateArray.length; l++) {
        if (array[l][0] == distinctDate[d] && array[l][1] == distinctLoc[i]) {

          if (chkDate == 0) {
            message = message + "\n\n" + distinctDate[d].toLocaleDateString('ru-RU',{ day: '2-digit', month: '2-digit' })
            chkDate = 1
          }

          message = message + "\n" + array[l][4] + " @" + array[l][5] + "\n" + array[l][2] + ": " + array[l][3]

        }
      }
    }
}
MESSAGE.getRange('B6').setValue(message)
Logger.log(message)
}
