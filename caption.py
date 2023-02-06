from cgitb import text
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import math
import pandas as pd
import CaptionData

#SETUP DATA
df = pd.read_excel("./data/data.xlsx", sheet_name="Sheet1")
captions = []
for row in df.values:
  captions.append(CaptionData.CaptionData(row[0], row[1], row[2], row[3], row[4]))

#CREATE PAGE
file_name = "captions.pdf"
file_path = os.path.join(os.getcwd(), "out", file_name)
page = canvas.Canvas(file_path, pagesize=landscape(A4), bottomup=False)

#SET FONT
pdfmetrics.registerFont(TTFont("REGULAR", "./data/GenShinGothic-Regular.ttf"))
pdfmetrics.registerFont(TTFont("MEDIUM", "./data/GenShinGothic-Medium.ttf"))

#DRAW
vw, vh = landscape(A4)
width = 100*mm
height = 55*mm
row_count = math.floor(vh/height)
column_count = math.floor(vw/width)
page_count = math.ceil(len(captions)/(row_count*column_count))

for p in range(page_count):
  for i in range(row_count):
    for j in range(column_count):
      #DATA
      index = row_count*column_count*p + column_count*i + j
      if index >= len(captions): break
      data = captions[index]

      #RECTANGLE
      origin_x = (vw-width*column_count)/2+width*j
      origin_y = (vh-height*row_count)/2+height*i
      page.rect(origin_x, origin_y, width, height)

      #TITLE
      margin = 8*mm
      page.setFont("MEDIUM", 19)
      title = data.title
      if len(title) >= 13:
        page.setFontSize(16)
      if len(title) >= 16:
        page.setFontSize(14)
      page.drawString(origin_x+margin, origin_y+margin+6*mm, title)

      #NAME
      name = data.name
      page.setFont("REGULAR", 15)
      page.drawString(origin_x+margin, origin_y+margin+19*mm, name)

      #University
      university = data.university
      page.setFont("REGULAR", 14)
      page.drawString(origin_x+margin+47*mm, origin_y+margin+19*mm, university)

      #MATERIAL
      material = data.material
      page.setFont("REGULAR", 12)
      if len(material) >= 20:
        page.setFontSize(10)
      page.drawRightString(origin_x+width-margin, origin_y+height-margin-10*mm, material)

      #SIZE
      size = data.size
      page.setFont("REGULAR", 12)
      page.drawRightString(origin_x+width-margin, origin_y+height-margin, size)
  page.showPage()

page.save()
