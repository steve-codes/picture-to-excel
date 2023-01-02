from PIL import Image
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

def main(): 
    #load the image
    i = Image.open(r'picture.jpg', 'r')

    #find the optimal image size to print in Excel (Excel only allows 65475 formatted cells)
    i = find_size(i)

    #create the Excel workbook and a new sheet within it
    wb = xlsxwriter.Workbook(r'Picture.xlsx')
    ws = wb.add_worksheet()

    #get the width and height of picture and load the pixels
    width, height = i.size
    pixels = i.load() 

    #for evey pixel in the width
    for x in range(width):
        #get the current column letter and set the width of the column to 20 pixels
        letter = xl_col_to_name(x)
        ws.set_column_pixels(letter + ":" + letter, width=20)
        #for evey pixel in the height
        for y in range(height): 
            #set the row height to 20 pixels 
            ws.set_row_pixels(y, 20)

            #grab the pixel information
            cpixel = pixels[x, y]

            #write the RGB values of a given pixel to the corresponding Excel file
            ws.write(xl_col_to_name(x) + str(y+1), '', wb.add_format({'bg_color': getHex(cpixel)}))
    #close the workbook
    wb.close()

#returns the hex RGB representation for inputting in Excel
def getHex(numlist):
    return "#" + hexFormat(numlist[0]) + hexFormat(numlist[1]) + hexFormat(numlist[2])

#strip away the 0x______ left hand side part of the hex code
def hexStrip(hexinput):
    return str(hex(hexinput)).split('x')[1]

#add a leading zero to the a given hex code (if not present)
def addZero(hexinput):
    if len(hexinput) == 1:
        return "0" + str(hexinput)
    return hexinput

#format the hex string
def hexFormat(hexinput):
    tmp = hexStrip(hexinput)
    tmp = addZero(tmp)
    return tmp

#recursive method for finding the optimal size for an image for Excel drawing    
def find_size(img):
    #get width and height 
    width, height = img.size
    
    #load the pixels
    pixels = img.load() 

    #dict for storing all RGB values
    dict= {}
    for x in range(width):
        for y in range(height):
            cpixel = pixels[x, y]
            dict[getHex(cpixel)] = ''
    
    #prints the width, height and length of the dict (total number of fromatting)
    print(width, height, len(dict))
    
    #if the length of the dict is greater than 65475, then reduce the image width and length by half
    if len(dict) > 65475:
        img = img.resize((int(width/2), int(height/2))) 
        return find_size(img)
    return img

if __name__ == "__main__":
    main()