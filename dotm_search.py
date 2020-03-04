#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import zipfile
import glob
import os


"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Cedric Mulvihill, Doug"




def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", help="add a directory")
    parser.add_argument("character", help="Character to search for")
    args = parser.parse_args()
    character = args.character
    directory = args.dir
    if directory:
        search_text(character, directory)
    

def search_text(character, path):
    files_matched = 0
    files_searched = 0
    os.chdir(path)
    files = glob.glob("*.dotm")
    for file in files:
        files_searched += 1
        with zipfile.ZipFile(file, 'r') as zfile:
            with zfile.open('word/document.xml', 'r') as read_file:
                for text in read_file:
                    if character in text:
                        files_matched += 1
    print("Number of Files Searched: " + str(files_searched))
    print("Number of Files Matched: " + str(files_matched))





if __name__ == '__main__':
    main()
