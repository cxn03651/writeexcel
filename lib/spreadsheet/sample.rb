require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

bp=1
    workbook  = Workbook.new('test.xls')
    worksheet = workbook.add_worksheet
    formula      = '=$E$3:$E$6'

    caption    = " \tData validation: _pack_dv_formula('#{formula}')"
    bytes      = %w(
                    09 00 0C 00 25 02 00 05 00 04 00 04 00
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(worksheet.pack_dv_formula(formula))
    p result
    p target

