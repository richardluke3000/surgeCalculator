
# for getting cmd arguments
import sys


# file location
file = sys.argv[1]

workbook = load_workbook( filename= file )

print (workbook.sheetnames)

