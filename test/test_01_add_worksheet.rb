#########################################
# test_01_add_worksheet.rb
#
# Tests for valid worksheet name handling.
#########################################
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

class TC_add_worksheet < Test::Unit::TestCase

   def test_ascii
      workbook = Excel.new
      assert_instance_of(Excel, workbook)
      worksheet1 = workbook.add_worksheet
      worksheet2 = workbook.add_worksheet
      worksheet3 = workbook.add_worksheet('sheet3')
      worksheet4 = workbook.add_worksheet('sheetz')
   end

end
