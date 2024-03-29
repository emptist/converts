pptxgen = require 'pptxgenjs'
pres = new pptxgen()

pres.author = 'ABC'
pres.company = 'SLL'
pres.revision = '15'
pres.subject = 'Annual Report'
pres.title = 'PptxGenJS Sample Presentation'

#pres.layout = 'LAYOUT_16x9'
pptx = pres

pptx.layout = "LAYOUT_WIDE"

pptx.defineSlideMaster({
    title: "MASTER_SLIDE",
    background: { color: "FFFFFF" },
    objects: [
        #{ line: { x: 3.5, y: 1.0, w: "100%", line: { color: "0088CC", width: 5 } } },
        { rect: { x: 0.0, y: 0.3, w: "100%", h: 0.75, fill: { color: "F1F1F1" } } },
        #{ text: { text: "Status Report", options: { x: 3.0, y: 0.3, w: 5.5, h: 0.75 } } },
        #{ image: { x: 11.3, y: "90%", w: 1.6, h: 0.5, path: "images/lotus001.jpeg" } },
        { placeholder:{
          options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 12, h: 5.25 },
          text: "(custom placeholder text!)",
        }}
        { placeholder:{
          options: { name: "title", type: "title", x: 0.3, y: 0.3, w:10, h: 0.75 },
          text: "(title placeholder text!)",
        }}
    ],
    slideNumber: { x: "90%", y: "90%" },
})

pptx.defineSlideMaster({
    title: "SECTION_MASTER_SLIDE",
    background: { color: "FFFFFF" },
    objects: [
        #{ line: { x: 3.5, y: 1.0, w: "100%", line: { color: "0088CC", width: 5 } } },
        { rect: { x: 0.0, y: 0.3, w: "100%", h: 0.75, fill: { color: "F1F1F1" } } },
        #{ text: { text: "Status Report", options: { x: 3.0, y: 0.3, w: 5.5, h: 0.75 } } },
        { image: { x: 11.3, y: "90%", w: 1.6, h: 0.5, path: "images/lotus001.jpeg" } },
        { placeholder:{
          options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 12, h: 5.25 },
          text: "(custom placeholder text!)",
        }}
        { placeholder:{
          options: { name: "title", type: "title", x: 0.3, y: 0.3, w:10, h: 0.75 },
          text: "(title placeholder text!)",
        }}
    ],
    slideNumber: { x: "90%", y: "90%" },
})

slide = pptx.addSlide({ masterName: "MASTER_SLIDE" })
slide.addText("How To Create PowerPoint Presentations with JavaScript", {placeholder:"title"})
slide.addText("This is easy", {placeholder:"body"})

#// Define new layout for the Presentation
#pptx.defineLayout({ name:'A3', width:16.5, height:11.7 })
#// Set presentation to use new layout
#pptx.layout = 'A3'
#slide = pres.addSlide("TITLE_SLIDE")

slide = pres.addSlide({ masterName: "MASTER_SLIDE" })

slide.background = { color: "F1F1F1" }  # Solid color
#slide.background = { color: "FF3399", transparency: 90 }  # hex fill color with transparency of 50%
#slide.background = { data: "image/pngbase64,ABC[...]123" }  # image: base64 data
#slide.background = { path: "https://some.url/image.jpg" }  # image: url

slide.color = "696969"  # Set slide default font color

#// EX: Add a Slide Number at a given location
#slide.slideNumber = { x: 1.0, y: "90%" }

#// EX: Styled Slide Numbers
#slide.slideNumber = { x: "95%", y: "95%", fontFace: "Courier", fontSize: 32, color: "FF3399" }

dataChartAreaLine = [
    {
        name: "Actual Sales",
        labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        values: [1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123, 15121],
    },
    {
        name: "Projected Sales",
        labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        values: [1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123, 12121],
    },
]

slide.addChart(pres.ChartType.radar, dataChartAreaLine, { x: 1, y: 1, w: 8, h: 4 })


#// STEP 1: Create a section
pptx.addSection({ title: "Tables" })

#// STEP 2: Provide section title to a slide that you want in corresponding section
slide = pptx.addSlide({ sectionTitle: "Tables" ,order: 2, masterName: "SECTION_MASTER_SLIDE"})

slide.addText("This slide is in the Tables section!", { x: 1.5, y: 1.5, fontSize: 18, color: "363636" })

slide = pptx.addSlide({ sectionTitle: "Tables" ,order: 3, masterName: "SECTION_MASTER_SLIDE"})

# TABLE 1: Single-row table
rows = [["Cell 1", "Cell 2", "Cell 3"]]
slide.addTable(rows, { w: 9 })

# TABLE 2: Multi-row table
# - each row's array element is an array of cells
slide = pptx.addSlide({ sectionTitle: "Tables" ,order: 5, masterName: "SECTION_MASTER_SLIDE"})
rows = [["A1", "B1", "C1"]]
slide.addTable(rows, { y:5, align: "left", fontFace: "Arial" })

# TABLE 3: Formatting at a cell level
# - use this to selectively override the table's cell options
rows = [
    [
        { text: "Top Lft", options: { align: "left", fontFace: "Arial" } },
        { text: "Top Ctr", options: { align: "center", fontFace: "Verdana" } },
        { text: "Top Rgt", options: { align: "right", fontFace: "Courier" } },
    ],
]
slide.addTable(rows, { y:3 , w: 9, rowH: 1, align: "left", fontFace: "Arial" })

#pptx.writeFile({ fileName: "Section Sample.pptx" })


#// For simple cases, you can omit `then`
# pptx.writeFile({ fileName: 'Browser-PowerPoint-Demo.pptx' })

#// Using Promise to determine when the file has actually completed generating
pptx.writeFile({ fileName: 'Gen-PowerPoint-Demo.pptx' })
    .then((fileName) -> 
        console.log("created file:#{fileName}")
    )