#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "jayaimzzz"

import re
import os
from zipfile import ZipFile
from lxml import etree

def main():
    path = "./dotm_files/"
    files = os.listdir(path)
    for file in files:
        if file.endswith("dotm"):
            print file
            with ZipFile(path + file, "r") as zip:
                # zip.printdir()
                print zip.
                # print dir(zip)
                # text_start_tag_locations = [l.start() for l in re.finditer("</w:t>", xml)]
                # i = xml.find("<w:t>")
                # print text_start_tag_locations
                # print xml
    

if __name__ == '__main__':
    main()
