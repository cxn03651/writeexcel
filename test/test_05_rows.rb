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
    @test_file           = "temp_test_file.xls"
    @workbook            = Excel.new(@test_file)
    @workbook.compatibility_mode(1)
  end

  def test_1
    row  = 1;
    col1 = 0;
    col2 = 0;
    worksheet = @workbook.add_worksheet
    worksheet.set_row(row, 15)
    push @tests,    [
      " \tset_row(): row = $row, col1 = $col1, col2 = $col2",
      {
        col_min => 0,
        col_max => 0,
      }
    ];
  end
end
