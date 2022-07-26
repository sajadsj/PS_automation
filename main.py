from photoshop import Session
import pandas as pd
# can also index sheet by name or fetch all sheets
df = pd.read_excel('C:/Users/Sajad/Desktop/names.xlsx')
mylist = df.values.tolist()

with Session() as ps:
    doc = ps.active_document
    # Add a solid color.
    textColor = ps.SolidColor()
    textColor.rgb.red = 0
    textColor.rgb.green = 0
    textColor.rgb.blue = 0
    options = ps.JPEGSaveOptions(quality=10)

    #  choose layer
    ps.echo(doc.activeLayer.name)
    ps.TextItem(doc.activeLayer.textItem)
    for item in df.values:
        student_name = str(item[0])
        doc.activeLayer.textItem.contents = student_name
        jpg_file = ("C:/Users/Sajad/Desktop/sut/"+student_name+".jpg")
        print(jpg_file)
        doc.saveAs(jpg_file, options, asCopy=True)
