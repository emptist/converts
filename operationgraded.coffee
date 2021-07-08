e2j = require 'convert-excel-to-json'
fs = require 'fs'
pptxgen = require 'pptxgenjs'
xlsx = require 'json-as-xlsx'

sourceFile = 'E:\projects\沧州中心医院量化项目\后期医院补报的资料\0702报艾力彼三四级手术占比2019-2020.xlsx'
jsonfilename = './手术占比.json'
outfilename = '沧州中心医院19'
pptname = '沧州中心医院专科报告19.pptx' 


readOpts = {
  sourceFile:sourceFile
  header: {rows: 2}
  sheets: ['二级专科']   #['三级专科','二级专科']
  #range: 'A6:Z14'
  columnToKey: {
    A:'心血管内科'	 
    B:'呼吸内科'	
    C:'消化内科'	
    D:'神经内科'	
    E:'肾内科'	
    F:'内分泌科'	
    G:'血液科'	
    H:'感染科'	
    I:'普通外科'	
    J:'胃肠外科'	
    k:'肝胆外科'	
    L:'乳腺外科'	
    M:'骨科'	
    N:'泌尿外科'	
    O:'神经外科'	
    P:'妇科'	
    Q:'产科'	
    R:'儿科'	
    S:'儿外科'	
    T:'新生儿科'	
    U:'眼科'	
    V:'口腔科'	
    W:'急诊科'	
    X:'重症医学科'	
    Y:'肿瘤内科'	
    Z:'老年内科'	
    #康复科	心脏外科	胸外科	烧伤整形外科	疼痛科	耳鼻喉科	皮肤科	风湿免疫科	血管外科	整形外科	介入科
  }
}

if fs.existsSync jsonfilename
  content = require jsonfilename
  #console.log content 
  
  arr = content[['二级专科']
  #console.log arr
  # 未完待续
  
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
  # console.log e2j 
  result = e2j readOpts
  
  jsonContent = JSON.stringify(result)

  fs.writeFile jsonfilename, jsonContent, 'utf8', (err) ->
    if err? 
      console.log(err)
    else
      console.log "json saved at #{Date()}"

#readToJson()