cej = require 'convert-excel-to-json'
fs = require 'fs'

# console.log cej 
result = cej {
  sourceFile:'/Users/jk/Downloads/艾力彼助理/jk/手术明细.xlsx'
  header: {rows: 1}
  sheets: ['Sheet1']
  #range: 'A2:H5'
  columnToKey: {
    A: '科室名称'	
    B: '科室诊断'	
    C: '麻醉级别'	
    D: '手术者'	
    E: '一助'	
    F: '二助'	
    G: '三助'	
    H: '四助'	
    I: '麻醉方式'	
    J: '麻醉医生'	
    K: '麻醉医生2'	
    L: '麻醉医生3'	
    M: '手术结束时间'	
    N: '手术开始时间'	
    O: '手术名称'	
    P: '手术时间分钟'	
    Q: '上台医师人数'
  }
}

jsonContent = JSON.stringify(result)

fs.writeFile 'operations.json', jsonContent, 'utf8', (err) ->
  if err? 
    console.log(err)
  else
    console.log 'json saved'