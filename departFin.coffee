cej = require 'convert-excel-to-json'
fs = require 'fs'
xlsx = require 'json-as-xlsx'

sourceFile = '/Users/jk/Downloads/建水2018-2020科室收入报表/2018-2020科室收入报表/2018年科室收入.xlsx'
jsonfilename = './departFin.json'
outfilename = 'departFin'
 
readOpts = {
  sourceFile:sourceFile
  header: {rows: 6}
  sheets: ['Sheet1']
  range: 'A6:Z14'
  columnToKey: {
    A: '科室名称'	
    C: '收入合计'	
    D: '门诊收入合计'	
    G: '门诊检查收入'	
    J: '门诊化验收入'
    K: '门诊耗材收入'	
    L: '门诊药品收入'	
    N: '住院收入合计'	
    Q: '住院检查收入'
    T: '住院化验收入'
    V: '住院耗材收入'
    W: '住院药品收入'
  }
}

if fs.existsSync jsonfilename
  content = require jsonfilename
  console.log content 
  
  arr = content.Sheet1
  console.log arr

  for each in arr
    if /门诊$/.test(each.科室名称)
      {门诊收入合计,门诊检查收入,门诊耗材收入,门诊药品收入,门诊化验收入} = each
      each.netIncome = 门诊收入合计 - (门诊检查收入+门诊耗材收入+门诊药品收入+门诊化验收入) 
    else
      {住院收入合计,住院检查收入,住院耗材收入,住院药品收入,住院化验收入} = each
      each.netIncome = 住院收入合计 - (住院检查收入+住院耗材收入+住院药品收入+住院化验收入)
    
    console.log each
  console.log arr
  
  unless fs.existsSync "#{outfilename}.xlsx"

    data = [
      {
        sheet: 'net income'
        columns: [
          {label:'科室名', value:'科室名称'}
          {label:'收入合计', value: '收入合计'}
          {label:'医疗服务收入', value: 'netIncome'}
          {label:'门诊检查收入', value: '门诊检查收入'}
          {label:'住院药品收入', value: '住院药品收入'}
        ]
        content: arr 
      }
    ]
    settings = {
      fileName: outfilename
      extraLength: 3
      writeOptions: {}
    }
    xlsx(data, settings)

else
  readToJson()


readToJson = () ->
  # console.log cej 
  result = cej readOpts
  
  jsonContent = JSON.stringify(result)

  fs.writeFile jsonfilename, jsonContent, 'utf8', (err) ->
    if err? 
      console.log(err)
    else
      console.log "json saved at #{Date()}"

#readToJson()