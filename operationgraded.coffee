e2j = require 'convert-excel-to-json'
fs = require 'fs'
pptxgen = require 'pptxgenjs'
xlsx = require 'json-as-xlsx'

sourceFile = './0702报艾力彼三四级手术占比2019-2020.xlsx'
jsonfilename = './手术占比.json'
pptname = './手术占比雷达图.pptx' 




readToJson = (readOpts) ->
  # console.log e2j 
  result = e2j readOpts
  
  jsonContent = JSON.stringify(result)

  fs.writeFile jsonfilename, jsonContent, 'utf8', (err) ->
    if err? 
      console.log(err)
    else
      console.log "json saved at #{Date()}"

readOpts = {
  sourceFile:sourceFile
  header: {rows: 2}
  sheets: ['三级专科','二级专科']  #['二级专科']
  #range: 'A6:Z14'
  columnToKey: {
    '*':'{{columnHeader}}'
    #A:"{{A2}}", B:"{{B2}}", C:"{{C2}}"
  }
}
###
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
    AA:'康复科'	
    AB:'心脏外科'	
    AC:'胸外科'	
    AD:'烧伤整形外科'	
    AE:'疼痛科'	
    AF:'耳鼻喉科'	
    AG:'皮肤科'	
    AH:'风湿免疫科'	
    AI:'血管外科'	
    AJ:'整形外科'	
    AK:'介入科'
  }
}

###
createPPT = (arr) ->
  labels = []
  values = []
  lvs = ({label: key, value: value} for key, value of obj for obj, index in arr)
  console.log "will create ppt here: ", lvs

  pres = new pptxgen()
  slide = pres.addSlide("TITLE_SLIDE")

  slide = pres.addSlide()

  #slide.background = { color: "F1F1F1" }  # hex fill color with transparency of 50%
  #slide.background = { data: "image/png;base64,ABC[...]123" }  # image: base64 data
  #slide.background = { path: "https://some.url/image.jpg" }  # image: url

  #slide.color = "696969"  # Set slide default font color

  # EX: Styled Slide Numbers
  slide.slideNumber = { x: "95%", y: "95%", fontFace: "Courier", fontSize: 32, color: "FF3399" }
  chartDataArray = [
    {
      name: "手术占比雷达图",
      labels: any.label for any in lvs[0],
      values: any.value for any in lvs[0]
    }
  ] 
      

  slide.addChart(pres.ChartType.radar, chartDataArray, { 
    x: 0, y: "50%", w: '45%', h: "50%" 
    showLegend: true, legendPos: "b"
  })
  slide.addChart(pres.ChartType.bar, chartDataArray, { 
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


if not fs.existsSync jsonfilename
  content = require jsonfilename
  #console.log content 
  
  arr = content['二级专科']
  #console.log arr
  
  unless fs.existsSync pptname
    createPPT(arr)
  
else
  readToJson(readOpts)

