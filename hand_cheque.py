from xlwt import Workbook
from    xlrd import open_workbook
import openpyxl
import argparse
import cv2
import numpy as np
import webcolors 
from PIL import ImageFont, ImageDraw, Image
import pytesseract
from pytesseract import image_to_string

image1 = input("Give the image name with extension : ")
image = cv2.imread(image1, cv2.IMREAD_COLOR)

print("\t\t\t=================================\n")
print("\t\t\tFEATURE EXTRACTION FROM THE IMAGE\n")
print("\t\t\t=================================\n")

img = cv2.imread(image1, cv2.IMREAD_COLOR)           # rgb
alpha_img = cv2.imread(image1, cv2.IMREAD_UNCHANGED) # rgba
gray_img = cv2.imread(image1, cv2.IMREAD_GRAYSCALE)  # grayscale

print ("\nProperties of the Image")
print ("**************************\n\n")
print ('RGB shape: ', img.shape)
print ("-------------------------\n")
print ('ARGB shape:', alpha_img.shape)
print ("-------------------------\n")
print ('Gray shape:', gray_img.shape)
print ("-------------------------\n")
print ('img.dtype: ', img.dtype)
print ("-------------------------\n")
print( 'img.size: ', img.size)
print ("-------------------------\n")

#Extracting for RGB
rgb=img[0][0]
#print(rgb)
r=rgb[0]
g=rgb[1]
b=rgb[2]
#print(r)
#print(g)
#print(b)

def closest_colour(requested_colour):
        min_colours = {}
        for key, name in webcolors.css3_hex_to_names.items():
                r_c, g_c, b_c = webcolors.hex_to_rgb(key)
                rd = (r_c - requested_colour[0]) ** 2
                gd = (g_c - requested_colour[1]) ** 2
                bd = (b_c - requested_colour[2]) ** 2
                min_colours[(rd + gd + bd)] = name
        return min_colours[min(min_colours.keys())]

def get_colour_name(requested_colour):
        try:
                closest_name = actual_name = webcolors.rgb_to_name(requested_colour)
        except ValueError:
                closest_name = closest_colour(requested_colour)
                #actual_name = None
        return    closest_name

requested_colour = (r,g,b)
closest_name = get_colour_name(requested_colour)

print("\nClosest colour name of the image :", closest_name)

# Open our image

im = Image.open(image1)
 
# Convert our image to RGB
rgb_im = im.convert('RGB')
 
# Use the .size object to retrieve a tuple contain (width,height) of the image
# and assign them to width and height variables
width = rgb_im.size[0]
height = rgb_im.size[1]
wid=str(width)
hgt=str(height)
print ("\nWidth and height of the image")
print("=================================")
print("\nWidth = " + wid + " pixels")
print("Height = " + hgt + " pixels")

print('--- Start recognize text from image ---')
print("Extracting Text .....\nThe text in the image is\n")

# initialize the list of reference points and boolean indicating
# whether cropping is being performed or not
#payee
ref_point = []
cropping = False
xmin=89
ymin=113
xmax=389
ymax=167
def shape_selection(event, x, y, flags, param):
  # grab references to the global variables
  global ref_point, cropping, xmin, ymin, xmax, ymax  

  ref_point = [(xmin, ymin)]
  cropping = True
  ref_point.append((xmax, ymax))
  cropping = False
  # draw a rectangle around the region of interest
  cv2.rectangle(image, ref_point[0], ref_point[1], (0, 255, 0), 2)
  cv2.imshow("image", image)
 
clone = image.copy()
cv2.namedWindow("image")
cv2.setMouseCallback("image", shape_selection)

# keep looping until the 'q' key is pressed
while True:
  # display the image and wait for a key press
  cv2.imshow("image", image)
  key = cv2.waitKey(1) & 0xFF

  # if the 'r' key is pressed, reset the cropping region
  if key == ord("r"):
    image = clone.copy()
    xmin=0
    ymin=0
    xmax=0
    ymax=0

  # if the 'c' key is pressed, break from the loop
  elif key == ord("c"):
    break

# if there are two reference points, then crop the region of interest
# from the image and display it
if len(ref_point) == 2:
  crop_img = clone[ref_point[0][1]:ref_point[1][1], ref_point[0][0]:ref_point[1][0]]
  #cv2.imshow("crop_img", crop_img)
  cv2.imwrite("hcpayee.png", crop_img)
  cv2.waitKey(0)

pytesseract.pytesseract.tesseract_cmd="C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
im=Image.open("G:/7 SEM MATERIALS/CAPSTONE/hcpayee.png")
payee=pytesseract.image_to_string(im,lang='eng')
print('Payee : ' + payee)

#Amount in words
ref_point = []
cropping = False
xmin=155
ymin=165
xmax=602
ymax=228
def shape_selection1(event, x, y, flags, param):
  # grab references to the global variables
  global ref_point, cropping, xmin, ymin, xmax, ymax  

  ref_point = [(xmin, ymin)]
  cropping = True
  ref_point.append((xmax, ymax))
  cropping = False
  # draw a rectangle around the region of interest
  cv2.rectangle(image, ref_point[0], ref_point[1], (0, 255, 0), 2)
  cv2.imshow("image", image)
 
clone = image.copy()
cv2.namedWindow("image")
cv2.setMouseCallback("image", shape_selection1)

# keep looping until the 'q' key is pressed
while True:
  # display the image and wait for a key press
  cv2.imshow("image", image)
  key = cv2.waitKey(1) & 0xFF

  # if the 'r' key is pressed, reset the cropping region
  if key == ord("r"):
    image = clone.copy()
    xmin=0
    ymin=0
    xmax=0
    ymax=0

  # if the 'c' key is pressed, break from the loop
  elif key == ord("c"):
    break

# if there are two reference points, then crop the region of interest
# from the image and display it
if len(ref_point) == 2:
  crop_img = clone[ref_point[0][1]:ref_point[1][1], ref_point[0][0]:ref_point[1][0]]
  #cv2.imshow("crop_img", crop_img)
  cv2.imwrite("hcamt_words.png", crop_img)
  cv2.waitKey(0)

pytesseract.pytesseract.tesseract_cmd="C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
im=Image.open("G:/7 SEM MATERIALS/CAPSTONE/hcamt_words.png")
amt_words=pytesseract.image_to_string(im,lang='eng')
print('Amount : ' + amt_words)

#Amount in Rupees
ref_point = []
cropping = False
xmin=948
ymin=216
xmax=1037
ymax=260
def shape_selection2(event, x, y, flags, param):
  # grab references to the global variables
  global ref_point, cropping, xmin, ymin, xmax, ymax  

  ref_point = [(xmin, ymin)]
  cropping = True
  ref_point.append((xmax, ymax))
  cropping = False
  # draw a rectangle around the region of interest
  cv2.rectangle(image, ref_point[0], ref_point[1], (0, 255, 0), 2)
  cv2.imshow("image", image)
 
clone = image.copy()
cv2.namedWindow("image")
cv2.setMouseCallback("image", shape_selection2)

# keep looping until the 'q' key is pressed
while True:
  # display the image and wait for a key press
  cv2.imshow("image", image)
  key = cv2.waitKey(1) & 0xFF

  # if the 'r' key is pressed, reset the cropping region
  if key == ord("r"):
    image = clone.copy()
    xmin=0
    ymin=0
    xmax=0
    ymax=0

  # if the 'c' key is pressed, break from the loop
  elif key == ord("c"):
    break

# if there are two reference points, then crop the region of interest
# from the image and display it
if len(ref_point) == 2:
  crop_img = clone[ref_point[0][1]:ref_point[1][1], ref_point[0][0]:ref_point[1][0]]
  #cv2.imshow("crop_img", crop_img)
  cv2.imwrite("hcamt.png", crop_img)
  cv2.waitKey(0)

pytesseract.pytesseract.tesseract_cmd="C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
im=Image.open("G:/7 SEM MATERIALS/CAPSTONE/hcamt.png")
amt=pytesseract.image_to_string(im,lang='eng')
print('Amount : ' + amt)

#Account Number
ref_point = []
cropping = False
xmin=80#89
ymin=153#285
xmax=186#332
ymax=174#315
def shape_selection3(event, x, y, flags, param):
  # grab references to the global variables
  global ref_point, cropping, xmin, ymin, xmax, ymax  

  ref_point = [(xmin, ymin)]
  cropping = True
  ref_point.append((xmax, ymax))
  cropping = False
  # draw a rectangle around the region of interest
  cv2.rectangle(image, ref_point[0], ref_point[1], (0, 255, 0), 2)
  cv2.imshow("image", image)
 
clone = image.copy()
cv2.namedWindow("image")
cv2.setMouseCallback("image", shape_selection3)

# keep looping until the 'q' key is pressed
while True:
  # display the image and wait for a key press
  cv2.imshow("image", image)
  key = cv2.waitKey(1) & 0xFF

  # if the 'r' key is pressed, reset the cropping region
  if key == ord("r"):
    image = clone.copy()
    xmin=0
    ymin=0
    xmax=0
    ymax=0

  # if the 'c' key is pressed, break from the loop
  elif key == ord("c"):
    break

# if there are two reference points, then crop the region of interest
# from the image and display it
if len(ref_point) == 2:
  crop_img = clone[ref_point[0][1]:ref_point[1][1], ref_point[0][0]:ref_point[1][0]]
  #cv2.imshow("crop_img", crop_img)
  cv2.imwrite("hcnum.png", crop_img)
  cv2.waitKey(0)

pytesseract.pytesseract.tesseract_cmd="C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
im=Image.open("G:/7 SEM MATERIALS/CAPSTONE/hcnum.png")
num=pytesseract.image_to_string(im,lang='eng')
print('A/C no : ' + num)

print("------ Done -------")

print("\n\nText from the image is extracted successfully !!!\n\n")

# Function to get the last RowCount in the Excel sheet , change the index of the sheet accordingly to get desired sheet.
def getDataColumn():
        #define the variables
        rowCount=0
        columnNumber=0
        wb = open_workbook('output.xlsx')
        ws = wb.sheet_by_index(0) 
        rowCount = ws.nrows
        rowCount+=1
        columnNumber=1 
	#print("The number of data in Excel file is ",rowcount)
        print("\n\nThe number of data in Excel file is ",rowCount)
        writedata(rowCount,columnNumber)

#Data to specified cells.
def writedata(rowNumber,columnNumber):
        book = openpyxl.load_workbook('output1.xlsx')
        sheet = book.get_sheet_by_name('Sheet1')
        data=[(str(image1),str(image1),hgt,wid,closest_name,str(rgb),payee,amt,amt_words,num)]
        #sheet.cell(row=rowNumber, column=columnNumber).value = 'kvs'
        # append all rows
        for row in data:
                sheet.append(row)
        book.save('output1.xlsx')
        print('\n\nThe above outputs are saved successfully in Excel File...\n')

getDataColumn()
