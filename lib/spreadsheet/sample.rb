require 'excel'
include Spreadsheet

   workbook  = Workbook.new('test.xls')
   worksheet = workbook.add_worksheet
bp=true
   result = worksheet.convert_date_time('2065-04-19T00:16:48.290')
   p result
   