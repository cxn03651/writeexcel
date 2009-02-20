#########################################
# tc_excel.rb
#
# Tests for the Excel class (excel.rb).
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
require "workbook"
require "worksheet"
require "excel"
require "ftools"

class TC_Excel < Test::Unit::TestCase
   def test_version
      assert_equal("0.3.5.1", Spreadsheet::Excel::VERSION)
   end
end
