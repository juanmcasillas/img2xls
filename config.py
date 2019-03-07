


class Config:
    def __init__(self):
        #https://support.microsoft.com/en-us/help/214123/description-of-how-column-widths-are-determined-in-excel
        #8.43
        # 1 inch -> 72 points
        # Calibri 11 (A) and Arial 10 (B)
      
        self.xls_row_height = 8.43     # excel row height
        self.xls_column_width = 1   # excel column width
        self.xls_zoom = 10  # small as possible # https://xlsxwriter.readthedocs.io/worksheet.html#set_zoom
        self.xls_gridlines = 0 # show all:0 , remove all: 2 # https://xlsxwriter.readthedocs.io/page_setup.html

        # constants
        self.xls_max_zoom = 80  

def setup_environment(args):

    # configure arguments
    config = Config()
    config.verbose = args.verbose
    config.inputfile = args.inputfile
    config.outputfile = args.outputfile
    config.noalpha = args.noalpha
    config.scale = args.scale
    config.pixelsize = args.pixelsize
    config.keepsmall = args.keepsmall
    return config
