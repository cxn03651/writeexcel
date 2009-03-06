require 'excel'
include Spreadsheet

   workbook  = Workbook.new('test.xls')
   worksheet = workbook.add_worksheet
bp=true
   result = worksheet.convert_date_time('0000-12-30T')
   p result
   