#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "andrewpchristensen701"

from zipfile import ZipFile 
import argparse
import os


def create_parser():
    parser = argparse.ArgumentParser(description='searches dotm files for text')
    parser.add_argument('--dir', help = 'directory to search', default='.') 
    parser.add_argument('text', help='text to search for')
    return parser

# directory

def main():
    parser = create_parser()
    args = parser.parse_args()
    print args

    fileList = os.listdir(args.dir)

    print fileList
    total_counter = 0
    match_counter = 0

    for file_0 in fileList: 
        if file_0.endswith('dotm'):
            with ZipFile(args.dir + '/' + file_0) as zip:

                with zip.open('word/document.xml') as doc:
                    total_counter += 1
                    lines = doc.read()
                    if lines.find(args.text) != -1:
                        print ("Match found in file: " +args.dir + '/'  + file_0 + '\n' 
                            + lines[lines.find(args.text)-40:lines.find(args.text)+40] + "\n")
                        match_counter += 1
                    
    print ("Total dotm files searched: " +  match_counter)
    print ("Total dotm files matched: " + total_counter)

                                
if __name__ == '__main__':
    main()