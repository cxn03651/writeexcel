require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

workbook  = Workbook.new('test.xls')
worksheet = workbook.add_worksheet
range = 'A6'
bp=1
worksheet.set_row(5, 6)
worksheet.set_row(6, 6)
worksheet.set_row(7, 6)
worksheet.set_row(8, 6)

data    = worksheet.substitute_cellref(range)
data    = worksheet.comment_params(data[0], data[1], 'Test')
data    = data[-1]
target  = %w(
00 00 10 F0 12 00
00 00 03 00 06 00 6A 00 01 00 69 00 06 00 F2 03
05 00 C4 00
).join(' ')

result  = unpack_record(worksheet.store_mso_client_anchor(3, *data))
p result

