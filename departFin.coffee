cej = require 'convert-excel-to-json'
fs = require 'fs'
pptxgen = require 'pptxgenjs'
xlsx = require 'json-as-xlsx'

sourceFile = '/Users/jk/Downloads/建水2018-2020科室收入报表/2018-2020科室收入报表/2018年科室收入.xlsx'
jsonfilename = './departFin.json'
outfilename = 'departFin'
pptname = 'departFin.pptx' 


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
  #console.log content 
  
  arr = content.Sheet1
  #console.log arr

  for each in arr
    if /门诊$/.test(each.科室名称)
      {门诊收入合计,门诊检查收入,门诊耗材收入,门诊药品收入,门诊化验收入} = each
      each.netIncome = 门诊收入合计 - (门诊检查收入+门诊耗材收入+门诊药品收入+门诊化验收入) 
    else
      {住院收入合计,住院检查收入,住院耗材收入,住院药品收入,住院化验收入} = each
      each.netIncome = 住院收入合计 - (住院检查收入+住院耗材收入+住院药品收入+住院化验收入)
    
    #console.log each
  #console.log arr
  
  if fs.existsSync pptname
    pres = new pptxgen()
    slide = pres.addSlide("TITLE_SLIDE")

    slide = pres.addSlide()

    #slide.background = { color: "F1F1F1" }  # hex fill color with transparency of 50%
    #slide.background = { data: "image/png;base64,ABC[...]123" }  # image: base64 data
    #slide.background = { path: "https://some.url/image.jpg" }  # image: url

    #slide.color = "696969"  # Set slide default font color

    # EX: Styled Slide Numbers
    slide.slideNumber = { x: "95%", y: "95%", fontFace: "Courier", fontSize: 32, color: "FF3399" }
    dataChartAreaLine = [
        {
            name: arr[0].科室名称,
            labels: ["医服收","收合","门收合","住收合"],
            values: [arr[0].医疗服务收入,arr[0].收入合计,arr[0].门诊收入合计,arr[0].住院收入合计]
        },
        {
            name: arr[1].科室名称,
            labels: ["医服收","收合","门收合","住收合"],
            values: [arr[1].医疗服务收入,arr[1].收入合计,arr[1].门诊收入合计,arr[1].住院收入合计]
        },
    ]

    slide.addChart(pres.ChartType.radar, dataChartAreaLine, { 
      x: 0, y: "50%", w: '45%', h: "50%" 
      showLegend: true, legendPos: "b"
    })
    slide.addChart(pres.ChartType.bar, dataChartAreaLine, { 
      x: 5, y: "50%", w: '45%', h: "50%" 
      showLegend: true, legendPos: "b"
      showTitle: true, title: "Bar Chart"
    })

    ###
    #// For simple cases, you can omit `then`
    pres.writeFile({ fileName: pptname})
    ###
    #// Using Promise to determine when the file has actually completed generating
    pres.writeFile({ fileName: pptname })
        .then((fileName) -> 
            console.log("created file:#{fileName} at #{Date()}")
        )

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