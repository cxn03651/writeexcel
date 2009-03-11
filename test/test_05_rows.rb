###############################################################################
#
# A test for Spreadsheet::WriteExcel.
#
# Check that max/min columns of the Excel ROW record are written correctly.
#
# reverse('©'), October 2007, John McNamara, jmcnamara@cpan.org
#
############################################################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i
  Dir.chdir("..") if base == "test"
  $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
  Dir.chdir("test") rescue nil
end

require "test/unit"
require "biffwriter"
require "olewriter"
require "format"
require "formula"
require "worksheet"
require "workbook"
require "excel"
include Spreadsheet

class TC_rows < Test::Unit::TestCase

  def setup
  end

  def test_1
    @test_file          = "temp_test_file.xls"
    workbook            = Excel.new(@test_file)
    workbook.compatibility_mode(1)
    @tests               = []
    
    # for test case 1
    row  = 1
    col1 = 0
    col2 = 0
    worksheet = workbook.add_worksheet
    worksheet.set_row(row, 15)
    @tests.push(
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
    @tests.push(
                 [
                    " \tset_row(): row = #{row}, col1 = #{col1}, col2 = #{col2}",
                    {
                      :col_min => 0,
                      :col_max => 1,
                    }
                 ]
              )


    row  = 3
    col1 = 0
    col2 = 1
    worksheet = workbook.add_worksheet
    worksheet.write($row, $col1, 'Test')
    worksheet.write($row, $col2, 'Test')
    @tests.push(
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
    # Read in the row records
    rows = []
  
    xlsfile = open(@test_file, "rb")
    while header = xlsfile.read(4)
      record, length = header.unpack('vv')
      data = xlsfile.read(length)
    
      #read the row records only
      next unless record == 0x0208
      col_min, col_max = data.unpack('x2 vv')
      
      rows.push(
        {
          :col_min => col_min,
          :col_max => col_max
        }
      )
    end
    (0 .. @tests.size - 1).each do |i|
      assert_equal(@tests[i][1], rows[i], @tests[i][0])
    end
  end

  def teardown
    
  end

end
