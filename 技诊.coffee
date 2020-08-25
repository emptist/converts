cej = require 'convert-excel-to-json'
fs = require 'fs'

# console.log cej 
result = cej {
  sourceFile:'/Users/jk/Downloads/艾力彼助理/jk/技诊2019_已处理.xlsx'
  header: {rows: 1}
  sheets: ['Sheet1']
  #range: 'A2:H5'
  columnToKey: {
    A: 'type'	
    B: 'itemID'
    C: 'cli_ward'	
    D: 'patientID'	
    E: 'start'	
    F: 'end'	
    G: 'minutes'	
    H: 'asked'	
    I: 'itemName'	
    J: 'cost'	
    K: 'result'	
    L: 'reporter'	
    M: 'machine'	
    N: 'amount'
  }
}

jsonContent = JSON.stringify(result)

fs.writeFile '技诊.json', jsonContent, 'utf8', (err) ->
  if err? 
    console.log(err)
  else
    console.log 'json saved'