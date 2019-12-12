# import the necessary packages
from xlwt import Workbook
from xlrd import open_workbook
import openpyxl
import argparse
import cv2
import numpy as np
import webcolors 
import os, os.path
from PIL import ImageFont, ImageDraw, Image
import pytesseract
from pytesseract import image_to_string

#debug info OpenCV version
print ("OpenCV version: " + cv2.__version__)

#image path and valid extensions
imageDir = "Data/" #specify your path here"
image_path_list = []
valid_image_extensions = [".jpg", ".jpeg", ".png", ".tif", ".tiff"] #specify your vald extensions here
valid_image_extensions = [item.lower() for item in valid_image_extensions]

#create a list all files in directory and
#append files with a vaild extention to image_path_list
for file in os.listdir(imageDir):
    extension = os.path.splitext(file)[1]
    if extension.lower() not in valid_image_extensions:
        continue
    image_path_list.append(os.path.join(imageDir, file))

#loop through image_path_list to open each image
for imagePath in image_path_list:
    #image = cv2.imread(imagePath)
    image = cv2.imread(imagePath, cv2.IMREAD_COLOR)
    #im=Image.open(image)
    #print(pytesseract.image_to_string(image,lang='eng'))
    print("\t\t\t=================================\n")
    print("\t\t\tFEATURE EXTRACTION FROM THE IMAGE\n")
    print("\t\t\t=================================\n")

    img = cv2.imread(imagePath, cv2.IMREAD_COLOR)                     # rgb
    alpha_img = cv2.imread(imagePath, cv2.IMREAD_UNCHANGED) # rgba
    gray_img = cv2.imread(imagePath, cv2.IMREAD_GRAYSCALE)    # grayscale

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

    im = Image.open(imagePath)
  
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
    print("\t\t\t=================================\n")
    print('--- Start recognize text from image ---')
    print("Extracting Text .....\nThe text in the image is\n")
    print("\t\t\t=================================\n")
    # initialize the list of reference points and boolean indicating
    # whether cropping is being performed or not
    #payee
    ref_point = []
    cropping = False
    xmin=57
    ymin=56
    xmax=363
    ymax=83
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
      cv2.imwrite("payee.png", crop_img)
      cv2.waitKey(0)

    pytesseract.pytesseract.tesseract_cmd="C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
    im=Image.open("G:/7 SEM MATERIALS/CAPSTONE/payee.png")
    payee=pytesseract.image_to_string(im,lang='eng')
    print('Payee : ' + payee)

    #Amount in words
    ref_point = []
    cropping = False
    xmin=95
    ymin=92
    xmax=478
    ymax=116
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
      cv2.imwrite("amt_words.png", crop_img)
      cv2.waitKey(0)

    pytesseract.pytesseract.tesseract_cmd="C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
    im=Image.open("G:/7 SEM MATERIALS/CAPSTONE/amt_words.png")
    amt_words=pytesseract.image_to_string(im,lang='eng')
    print('Amount : ' + amt_words)

#Amount in Rupees
    ref_point = []
    cropping = False
    xmin=524
    ymin=115
    xmax=637
    ymax=141
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
        
  # if the 'c' key is pressed, break from the loop
      elif key == ord("c"):
        break

# if there are two reference points, then crop the region of interest
# from the image and display it
    if len(ref_point) == 2:
      crop_img = clone[ref_point[0][1]:ref_point[1][1], ref_point[0][0]:ref_point[1][0]]
      #cv2.imshow("crop_img", crop_img)
      cv2.imwrite("amt.png", crop_img)
      cv2.waitKey(0)

    pytesseract.pytesseract.tesseract_cmd="C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
    im=Image.open("G:/7 SEM MATERIALS/CAPSTONE/amt.png")
    amt=pytesseract.image_to_string(im,lang='eng')
    print('Amount : ' + amt)

#Account Number
    ref_point = []
    cropping = False
    xmin=74
    ymin=155
    xmax=203
    ymax=172
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
        
  # if the 'c' key is pressed, break from the loop
      elif key == ord("c"):
        break

# if there are two reference points, then crop the region of interest
# from the image and display it
    if len(ref_point) == 2:
      crop_img = clone[ref_point[0][1]:ref_point[1][1], ref_point[0][0]:ref_point[1][0]]
      #cv2.imshow("crop_img", crop_img)
      cv2.imwrite("num.png", crop_img)
      cv2.waitKey(0)

    pytesseract.pytesseract.tesseract_cmd="C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
    im=Image.open("G:/7 SEM MATERIALS/CAPSTONE/num.png")
    num=pytesseract.image_to_string(im,lang='eng')
    print('A/C No.: ' + num)

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
            #print("\n\nThe number of data in Excel file is ",rowCount)
            writedata(rowCount,columnNumber)

#Data to specified cells.
    def writedata(rowNumber,columnNumber):
            book = openpyxl.load_workbook('output1.xlsx')
            sheet = book.get_sheet_by_name('Sheet1')
            data=[(str(imagePath),str(imagePath),hgt,wid,closest_name,str(rgb),payee,num,amt,amt_words)]
        #sheet.cell(row=rowNumber, column=columnNumber).value = 'kvs'
        # append all rows
            for row in data:
                    sheet.append(row)
            book.save('output1.xlsx')
            print('\n\nThe above outputs are saved successfully in Excel File...\n')

    getDataColumn()

    
    # display the image on screen with imshow()
    # after checking that it loaded
    if image is not None:
        cv2.imshow(imagePath, image)
    elif image is None:
        print ("Error loading: " + imagePath)
        #end this loop iteration and move on to next image
        continue
    
    # wait time in milliseconds
    # this is required to show the image
    # 0 = wait indefinitely
    # exit when escape key is pressed
    key = cv2.waitKey(0)
    if key == 27: # escape
        break

# close any open windows
cv2.destroyAllWindows()