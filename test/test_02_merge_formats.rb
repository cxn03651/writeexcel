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

class TC_merge_formats < Test::Unit::TestCase

   def test_some
      setup_test

      #Test 1
      assert_nothing_raised { @worksheet.write('A1',    'Test', @non_merged_format) }
      assert_nothing_raised { @worksheet.write('A3:B4', 'Test', @merged_format) }
   end

   def setup_test
      @test_file           = "temp_test_file.xls"
      @workbook            = Excel.new(@test_file)
      @worksheet           = @workbook.add_worksheet
      @merged_format       = @workbook.add_format(:bold => 1)
      @non_merged_format   = @workbook.add_format(:bold => 1)

      @worksheet.set_row(    5,    nil, @merged_format)
      @worksheet.set_column('G:G', nil, @merged_format)
   end

end
