require 'excel'
include Spreadsheet

   workbook  = Workbook.new('test.xls')
   worksheet = workbook.add_worksheet
   smiley    = [0x263a].pack('n')
   format = workbook.add_format
   formula = worksheet.store_formula('A1*3+50')
bp=true
   worksheet.repeat_formula(5, 3, formula, format, 'A1', 'A2')
   worksheet.write_formula(1,2,'1+2')
   formula = Formula.new(0)
   p = formula.parse('1+2')
   p formula.reverse(p)
      