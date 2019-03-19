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
import argparse
import sys
import time

def main(directory, search_text):
    # t0 = time.time()
    path = directory + "/" if directory != None else "./"
    files = os.listdir(path)
    files_matched = 0
    files_searched = 0
    len_of_search_text = len(search_text)
    print "Searching director {} for text '{}' ...".format(path,search_text)
    for fil in files:
        if fil.endswith("dotm"):
            files_searched += 1
            with ZipFile(path + fil, "r") as zip:
                with zip.open("word/document.xml", "r") as opened:
                    for text in opened:
                        if search_text in text:
                            files_matched += 1
                            print "Match found in file {}{}".format(path,fil)
                            len1 = len(text)
                            for i in range(len1):                           
                                if text[i:i + len_of_search_text] == search_text:
                                    start_i = 0 if i < 40 else i - 40
                                    end_i = len1 if i + 40 + len_of_search_text > len1 else i + 40 + len_of_search_text
                                    print "   ...{}...".format(text[start_i:end_i])
    print "Total dotm files searched: {}".format(files_searched)
    print "Total dotm files matched: {}".format(files_matched)
    # t1 = time.time()
    # print "Total time: {}".format(t1-t0)

if __name__ == '__main__':
    """Search a directory of dotm files for text"""
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir",help="directory to search")
    parser.add_argument("search_text", help="the text to search for")
    args = parser.parse_args()
    main(args.dir,args.search_text)
