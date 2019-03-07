from PIL import Image
import xlsxwriter
import os.path
import math
import datetime

class Converter:

    SHOW_INFO = 1

    def __init__(self, config):
        self.config = config
        
        # map some handy shortcuts
        self.verbose = config.verbose
        self.inputfile = config.inputfile
        self.outputfile = config.outputfile

    def ReadFile(self):
        try:
            self.image = Image.open(self.inputfile)

            if self.verbose >= Converter.SHOW_INFO:
                self.print_info()

        except Exception as e:
            print("can't open %s" % self.inputfile)
            raise e 

    def  ProcessFile(self):
        self.image_rgb = self.image.copy()
        self.image_rgb.convert('RGBA')

        if self.config.scale:
            self.image_rgb = self.scale(self.image_rgb, self.config.scale)

        if self.config.pixelsize:
            self.image_rgb = self.pixelize(self.image_rgb, self.config.pixelsize)
    
    def GenerateXLS(self):


        workbook = xlsxwriter.Workbook(self.outputfile)
        worksheet = workbook.add_worksheet(os.path.basename(self.inputfile))

        # set options (2 removes all the gridlines)

        self.set_properties(workbook)
        worksheet.hide_gridlines(self.config.xls_gridlines) 
        worksheet.set_zoom(self.config.xls_zoom) 
        formats = {}

        for x in range(self.image_rgb.size[0]):

            worksheet.set_column(x,x, self.config.xls_column_width)
            for y in range(self.image_rgb.size[1]):
                worksheet.set_row(y, self.config.xls_row_height)
                
                pixel_info = self.image_rgb.getpixel((x, y))

                r = pixel_info[0]
                g = pixel_info[1]
                b = pixel_info[2]
                a = 255 # alpha channel (255: fully transparent, 0: fully opaque)
                if self.image_rgb.mode == 'RGBA':
                    a = pixel_info[3]
                color_str = "#%02x%02x%02x%02x" % (r,g,b,a)
                cell_format = None

                if not color_str in formats.keys():
                    cell_format = workbook.add_format()

                    if a == 255 or self.config.noalpha == True:
                        color_str_without_alpha = color_str[:-2]
                        cell_format.set_bg_color(color_str_without_alpha)
                    else:
                        pass
                        
                    formats[color_str] = cell_format
                else:
                    cell_format = formats[color_str]
                
                worksheet.write_blank(y,x, None, cell_format)


        workbook.close()

    def set_properties(self, workbook):
        workbook.set_properties({
        'title':     self.inputfile,
        'subject':  'Img2XLS conversion',
        'author':   'Juan M. Casillas',
        'manager':  'Juan M. Casillas',
        'category': 'PIXEL Art',
        'keywords': 'PIXEL Art',
        'created':  datetime.datetime.now(),
        'comments': 'Created with Python, Img2XLS, PILLOW and XlsxWriter', 
        'hyperlink_base': 'https://github.com/juanmcasillas/img2xls'
    })

    def print_info(self):
        print("File: %s" % self.inputfile)
        print(" ->Width: %6d" % self.image.size[0])
        print(" ->Height: %5d" % self.image.size[1])
        print(" ->Format: %5s" % self.image.format)
        print(" ->Mode: %7s" % self.image.mode)

    def scale(self, img, percent):

        basewidth = img.size[0]* percent / 100.0

        wpercent = (basewidth/float(img.size[0]))
        hsize = int((float(img.size[1])*float(wpercent)))

        basewidth = int(basewidth)
        hsize = int(hsize)
        imgr = img.resize( (basewidth,hsize) , Image.ANTIALIAS)

        if self.verbose >= Converter.SHOW_INFO:
            print("Scaling image [(%d,%d) - %d%% -> (%d,%d)]" % (img.size[0],img.size[1],percent, basewidth, hsize))

        return(imgr)

    def pixelize(self, img, pixelsize=16):

        wpercent = (pixelsize/float(img.size[0]))
        hsize = int((float(img.size[1])*float(wpercent)))

        pixelsize_x = int(math.ceil(pixelsize))
        pixelsize_y = int(math.ceil(hsize))
        imgSmall = img.resize((pixelsize_x,pixelsize_y),resample=Image.BILINEAR)

        if self.config.keepsmall == False:
            # Scale back up using NEAREST to original size
            imgPixel = imgSmall.resize(img.size,Image.NEAREST)
        else:
            # reescale excel size to create a big picture
            imgPixel = imgSmall
            self.config.xls_zoom = self.config.xls_max_zoom

        if self.verbose >= Converter.SHOW_INFO:
            print("Pixelizing image. Pixel %d grid (%d,%d) for aspect ratio" % (pixelsize, pixelsize_x,pixelsize_y))

        return imgPixel