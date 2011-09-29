#!/usr/bin/env python2.6
'''
Sample program to extract the media contents of a docx file.
ALPHA!
'''
from docXtractor import *
import sys
import os
if __name__ == '__main__':        
    try:
        k = extractall(sys.argv[1])

        f = open('test.csv', 'w')
	for i in k:
	    f.write(i + '\n')
	f.close()

    except:
        print('Done.')
        exit()
