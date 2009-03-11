require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

bp=1
    test_file           = "temp_test_file.xls"
    workbook            = Excel.new(test_file)
    workbook.compatibility_mode(1)
    tests               = []
    
    # for test case 1
    row  = 1
    col1 = 0
    col2 = 0
    worksheet = workbook.add_worksheet
    worksheet.set_row(row, 15)
    tests.push(
                 [
                    " \tset_row(): row = #{row}, col1 = #{col1}, col2 = #{col2}",
                    {
                      :col_min => 0,
                      :col_max => 0,
                    }
                 ]
              )

    # for test case 2
    row  = 2
    col1 = 0
    col2 = 0
    worksheet = workbook.add_worksheet
    worksheet.write(row, col1, 'Test')
    worksheet.write(row, col2, 'Test')
    tests.push(
                 [
                    " \tset_row(): row = #{row}, col1 = #{col1}, col2 = #{col2}",
                    {
                      :col_min => 0,
                      :col_max => 1,
                    }
                 ]
              )

    # for test case 3
    row  = 3
    col1 = 0
    col2 = 1
    worksheet = workbook.add_worksheet
    worksheet.write(row, col1, 'Test')
    worksheet.write(row, col2, 'Test')
    tests.push(
                [
                    " \twrite():   row = $row, col1 = $col1, col2 = $col2",
                    {
                        :col_min => 0,
                        :col_max => 2,
                    }
                ]
            )


    workbook.biff_only  = 1
    workbook.close

    rows = []
  
    xlsfile = open(test_file, "rb")
    while header = xlsfile.read(4)
      record, length = header.unpack('vv')
      data = xlsfile.read(length)
    
      #read the row records only
      next unless record == 0x0208
      col_min, col_max = data.unpack('x2 vv')
print "record = #{record}, length = #{length}, col_min = #{col_min}, col_max = #{col_max}\n"
      
      rows.push(
        {
          :col_min => col_min,
          :col_max => col_max
        }
      )
    end

