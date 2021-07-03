pptxgen = require 'pptxgenjs'
pres = new pptxgen()

pres.author = 'ABC'
pres.company = 'SLL'
pres.revision = '15'
pres.subject = 'Annual Report'
pres.title = 'PptxGenJS Sample Presentation'

#pres.layout = 'LAYOUT_16x9'
pptx = pres

#// Define new layout for the Presentation
pptx.defineLayout({ name:'A3', width:16.5, height:11.7 })
#// Set presentation to use new layout
pptx.layout = 'A3'
slide = pres.addSlide("TITLE_SLIDE")
slide = pres.addSlide()

slide.background = { color: "F1F1F1" }  # Solid color
slide.background = { color: "FF3399", transparency: 50 }  # hex fill color with transparency of 50%
#slide.background = { data: "image/png;base64,ABC[...]123" }  # image: base64 data
#slide.background = { path: "https://some.url/image.jpg" }  # image: url

slide.color = "696969"  # Set slide default font color

#// EX: Add a Slide Number at a given location
#slide.slideNumber = { x: 1.0, y: "90%" }

#// EX: Styled Slide Numbers
slide.slideNumber = { x: 1.0, y: "95%", fontFace: "Courier", fontSize: 32, color: "CF0101" }

#// For simple cases, you can omit `then`
# pptx.writeFile({ fileName: 'Browser-PowerPoint-Demo.pptx' })

#// Using Promise to determine when the file has actually completed generating
pptx.writeFile({ fileName: 'Gen-PowerPoint-Demo.pptx' })
    .then((fileName) -> 
        console.log("created file:#{fileName}")
    )