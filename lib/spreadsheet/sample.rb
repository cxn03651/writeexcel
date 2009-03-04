require 'excel'
include Spreadsheet

wb = Excel.new('test.xls')
worksheet = wb.add_worksheet
p worksheet.get_name
worksheet1 = wb.add_worksheet('Sheet1')
