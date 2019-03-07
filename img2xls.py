#!/usr/bin/env python
#
# (c) 07/03/2019 Juan M. Casillas
# juanm.casillas@gmail.com
# 
# https://github.com/juanmcasillas/img2xls

import argparse 
from PIL import Image


import config
import converter


def parseArgs():
    parser = argparse.ArgumentParser()
    parser.add_argument("-v", "--verbose", help="increase output verbosity", action="count")
    parser.add_argument("-n", "--noalpha", help="Skip alpha channel", action="store_true")
    parser.add_argument("-s", "--scale", help="Percent scale (200 -> double, 50 -> half)", type=int)
    parser.add_argument("-p", "--pixelsize", help="Create a pixelsize efect of PxP pixels (e.g. 16x16).", type=int)
    parser.add_argument("-k", "--keepsmall", help="Don't resize again when pixelize", action="store_true")
    parser.add_argument("inputfile", help="The image file to be read")
    parser.add_argument("outputfile", help="output file. Generates an EXCEl compatible file")
    args = parser.parse_args()

    return args        

if __name__ == "__main__":

    args = parseArgs()
    config = config.setup_environment(args)
    conv = converter.Converter(config)
    conv.ReadFile()
    conv.ProcessFile()
    conv.GenerateXLS()