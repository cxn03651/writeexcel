require 'excel'
include Spreadsheet

   workbook          = Excel.new('test.xls')
   worksheet         = workbook.add_worksheet
   merged_format     = workbook.add_format(:bold => 1)
   non_merged_format = workbook.add_format(:bold => 1)
   worksheet.set_row(5, nil, merged_format)
   worksheet.set_column('G:G', nil, merged_format)
breakpoint = true
   worksheet.write('A1',    'Test', non_merged_format)
   worksheet.write('A3:B4', 'Test', merged_format)
   