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
require 'writeexcel'

class TC_add_worksheet < Test::Unit::TestCase

  def test_true
    assert(true)
  end

end
