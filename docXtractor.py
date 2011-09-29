#!/usr/bin/env python2.8
'''
Extract media contents from Office Open XML (docx) files
'''
from lxml import etree
import zipfile
import shutil
import re
import time
import os
import hashlib
from os.path import join
try:
    from PIL import Image
except ImportError:
    import Image

template_dir = join(os.path.dirname(__file__),'docx-template') 
if not os.path.isdir(template_dir):
    template_dir = join(os.path.dirname(__file__),'template') 

names = {
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
}

def opendocx(file):
    '''Open a docx file, return a document XML tree'''
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)
    return document

# Sample media extraction
def extracttest(file):
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml')
    document = etree.fromstring(xmlcontent)
    l = mydoc.namelist()
    for i in l:
        print i
        if i.find('word/media') > -1:
           mydoc.extract(i, 'extracted')
    return document

# Extract all media
def extractall(indir):

    ml = []
    ml.append('FILENAME, DESCRIPTION, IMAGE NAME, RELATIVE ID')
    #ml.append('FILENAME, IMAGE ID, DESCRIPTION, IMAGE NAME, RELATIVE ID')

    # Extract all media to a specified dir (with word subdirs) from cmdline
    for dirname, dirnames, filenames in os.walk(indir):
        for subdirname in dirnames:
            os.path.join(dirname, subdirname)
	for filename in filenames:

            print 'Processing ' + os.path.join(dirname, filename) + '...'
    	    mydoc = zipfile.ZipFile(os.path.join(dirname, filename))
            xmlcontent = mydoc.read('word/document.xml')
            document = etree.fromstring(xmlcontent)
            l = mydoc.namelist()
            curr_hash = ''
            for i in l:
                if i.find('word/media') > -1:
                     mydoc.extract(i, os.path.join(dirname, filename) + ' extracted')
	    
	    #get the initial relationship mapping
            relscontent = mydoc.read('word/document.xml')
	    reldoc = etree.fromstring(relscontent)            	
            for element in reldoc.iter():
		if 'graphicData' in element.tag:
                    mys = ''
		    for child in element.iter():
			if 'cNvPr' in child.tag:
		           attributes = child.attrib
		           if attributes.get("descr") == None:
		    	       continue
		           else:
		    	       mys = attributes.get("descr") + ',' + attributes.get("name")

		        if child.tag[-4:] == 'blip':
			    attributes = child.attrib
	
		            if attributes.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed") == None:
			        continue
			    else:
                                tmp = attributes.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
		  	    if len(mys) != 0:		
			        ml.append(filename + ',' + mys + ',' + tmp)

    return ml

