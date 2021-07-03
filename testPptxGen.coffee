pptxgen = require 'pptxgenjs'
pres = new pptxgen()

pres.author = 'ABC'
pres.company = 'SLL'
pres.revision = '15'
pres.subject = 'Annual Report'
pres.title = 'PptxGenJS Sample Presentation'

pres.layout = 'LAYOUT_16x9'
pptx = pres

#// Define new layout for the Presentation
#pptx.defineLayout({ name:'A3', width:16.5, height:11.7 })
#// Set presentation to use new layout
#pptx.layout = 'A3'
slide = pres.addSlide("TITLE_SLIDE")

slide = pres.addSlide()

slide.background = { color: "F1F1F1" }  # Solid color
#slide.background = { color: "FF3399", transparency: 90 }  # hex fill color with transparency of 50%
#slide.background = { data: "image/png;base64,ABC[...]123" }  # image: base64 data
#slide.background = { path: "https://some.url/image.jpg" }  # image: url

slide.color = "696969"  # Set slide default font color

#// EX: Add a Slide Number at a given location
#slide.slideNumber = { x: 1.0, y: "90%" }

#// EX: Styled Slide Numbers
slide.slideNumber = { x: "95%", y: "95%", fontFace: "Courier", fontSize: 32, color: "FF3399" }

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

slide.addChart(pres.ChartType.line, dataChartAreaLine, { x: 1, y: 1, w: 8, h: 4 })

#// For simple cases, you can omit `then`
# pptx.writeFile({ fileName: 'Browser-PowerPoint-Demo.pptx' })

#// Using Promise to determine when the file has actually completed generating
pptx.writeFile({ fileName: 'Gen-PowerPoint-Demo.pptx' })
    .then((fileName) -> 
        console.log("created file:#{fileName}")
    )