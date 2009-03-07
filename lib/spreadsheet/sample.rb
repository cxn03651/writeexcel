require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

bp=1
    workbook  = Workbook.new('test.xls')
    worksheet = workbook.add_worksheet
    data      = worksheet.comment_params(2,0,'Test')
    row       = data[0]
    col       = data[1]
    author    = data[4]
    encoding  = data[5]
    visible   = data[6]
    obj_id    = 1
    
    result = unpack_record(
        worksheet.store_note(row,col,obj_id,author,encoding,visible))
p result

