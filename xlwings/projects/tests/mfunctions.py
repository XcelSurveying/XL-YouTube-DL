import numpy as np
from xlwings import Workbook, Range

def rand_numbers():
    """ produces standard normally distributed random numbers with shape (n,n)"""
    wb = Workbook.caller()  # Creates a reference to the calling Excel file
    n = int(Range('Sheet1', 'B1').value)  # Write desired dimensions into Cell B1
    rand_num = np.random.randn(n, n)
    Range('Sheet1', 'C3').value = rand_num
	
def array():
	""" Test array input from Python to Excel"""
	wb = Workbook.caller() # Creates a reference to the calling excel file
	Range('M8').value = np.eye(5)
	A = np.array(Range('A1', asarray=True).table.value)
	Range('M20').value = A
	#A = int(Range('Sheet1', 'M3:O6')) # Range M3:O6
	#Range('A15').value = A
