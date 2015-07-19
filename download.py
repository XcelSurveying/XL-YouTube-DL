from __future__ import print_function, unicode_literals
import pafy
import sys
from xlwings import Workbook, Range

def downloadVidList():
        collection = ['Vsy1URDYK88']

        for url in collection:
                video = pafy.new(url)

                best = video.getbest()
                best.download(quiet=False)
	
def VideoDescription():

        url = Range('Sheet1', 'C3').value
        video = pafy.new(url)
        
        wb = Workbook.caller() # Creates a reference to the calling Excel file

        videoTitle = video.title
        
        Range('Sheet1', 'G3').value = videoTitle

        
