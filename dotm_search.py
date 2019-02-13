#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "jayaimzzz"

import os
from zipfile import ZipFile

def main():
    path = "./dotm_files/"
    files = os.listdir(path)
    for file in files:
        if file.endswith("dotm"):
            print file
            with ZipFile(path + file, "r") as zip:
                # zip.printdir()
                data = zip.read("word/document.xml")
                print data
    

if __name__ == '__main__':
    main()
