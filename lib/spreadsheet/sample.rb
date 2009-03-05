require 'excel'
include Spreadsheet

   workbook = Workbook.new('test.xls')
   worksheet = workbook.add_worksheet
bp = true
   worksheet.write_formula(1,2,'1+2')
   formula = Formula.new(0)
   p = formula.parse('1+2')
   p formula.reverse(p)
      