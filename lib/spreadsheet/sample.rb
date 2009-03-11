require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

bp=1
    test_file           = "temp_test_file.xls"
    workbook            = Excel.new(test_file)
    workbook.compatibility_mode(1)
    worksheet = workbook.add_worksheet
    range = 'A6'
    worksheet.set_row(5, 6)
    worksheet.set_row(6, 6)
    worksheet.set_row(7, 6)
    worksheet.set_row(8, 6)
    data    = worksheet.substitute_cellref(range)
    data    = worksheet.comment_params(data[0], data[1], 'Test')
    data    = $data[-1]

    workbook.biff_only  = 1
    workbook.close

