require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

bp=1
    workbook  = Workbook.new('test.xls')
    worksheet = workbook.add_worksheet
    count = 1
    for i in 1 .. count
      worksheet.write_comment(i -1, 0, 'aaa')
    end
    workbook.calc_mso_sizes
    target  = %w(
        EB 00 5A 00 0F 00 00 F0 52 00 00 00 00 00 06 F0
        18 00 00 00 02 04 00 00 02 00 00 00 02 00 00 00
        01 00 00 00 01 00 00 00 02 00 00 00 33 00 0B F0
        12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
        C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
        00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
    ).join(' ')
    result = unpack_record(workbook.add_mso_drawing_group)
p result

