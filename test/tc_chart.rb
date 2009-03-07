#####################################################
# tc_chart.rb
#
# Test suite for the Chart class (chart.rb)
#####################################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i
  Dir.chdir("..") if base == "test"
  $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
  Dir.chdir("test") rescue nil
end

require "test/unit"
require "biffwriter"
require "olewriter"
require "workbook"
require "worksheet"
require "format"
require 'formula'

class TC_Chart < Test::Unit::TestCase

  def test_
    assert(true)
  end

end
