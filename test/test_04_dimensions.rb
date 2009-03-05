###########################################################################
# test_02_merge_formats.rb
#
# Tests to ensure merge formats aren't used in non-merged cells and
# vice-versa. This is temporary feature to prevent users from inadvertently
# making this error.
#
# reverse('©'), April 2005, John McNamara, jmcnamara@cpan.org
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

class TC_dimensions < Test::Unit::TestCase

   def setup
      @test_file           = "temp_test_file.xls"
      @workbook            = Excel.new(@test_file)
      @worksheet           = @workbook.add_worksheet
      @format              = @workbook.add_format
      @dims                = ['row_min', 'row_max', 'col_min', 'col_max']
      @smiley              = [0x263a].pack('n')
   end

   def test_no_worksheet_cell_data
      data     = @worksheet.store_dimensions

      vals     = data.unpack('x4 VVvv')
      alist    = @dims.zip(vals)
      results  = Hash[*alist.flatten]

      alist    = @dims.zip([0, 0, 0, 0])
      expected = Hash[*alist.flatten]

      assert_equal(expected, results)
   end

   def test_data_in_cell_0_0
      @worksheet.write(0, 0, 'Test')
      data     = @worksheet.store_dimensions

      vals     = data.unpack('x4 VVvv')
      alist    = @dims.zip(vals)
      results  = Hash[*alist.flatten]

      alist    = @dims.zip([0, 1, 0, 1])
      expected = Hash[*alist.flatten]
      
      assert_equal(expected, results)
   end

end
