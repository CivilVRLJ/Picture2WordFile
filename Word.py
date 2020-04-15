from docx import Document
from docx.shared import Cm
import os
import numpy as np

title = "Erosion vs Height"

document = Document()
document.add_heading(str(title), 0)

areaList = os.listdir(str(title))
print("areaList: " + str(areaList))

if ".DS_Store" in areaList:
    areaList.remove(".DS_Store")

areaList = np.sort(areaList)

for area in areaList:

    listFiles = os.listdir(str(title) + "/" + str(area))
    listFiles = np.sort(listFiles)

    listImage = []
    for file in listFiles:
        print("file: " + str(file))
        listImage.append(file)

    document.add_heading(area, level=1)
    for pngName in listImage:
        document.add_picture(str(title) + "/" + str(area) + "/" + str(pngName), width=Cm(20))
        document.add_page_break()

document.save(str(title) + '.docx')
