#######################################################
# tc_example_match.rb
#
# Test suite for matching with xls file made by perl.
#######################################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i
  Dir.chdir("..") if base == "test"
  $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
  Dir.chdir("test") rescue nil
end

require "test/unit"
require "excel"
include Spreadsheet

class TC_example_match < Test::Unit::TestCase

  def setup
    @filename = "tc_example_match.xls"
  end

  def teardown
    File.delete(@filename) if File.exist?(@filename)
  end

  def test_a_simple
    xlsfile   = 'a_simple.xls'
    workbook  = Excel.new(xlsfile);
    worksheet = workbook.add_worksheet
    
    # The general syntax is write(row, column, token). Note that row and
    # column are zero indexed
    #
    
    # Write some text
    worksheet.write(0, 0,  "Hi Excel!")
    
    
    # Write some numbers
    worksheet.write(2, 0,  3)          # Writes 3
    worksheet.write(3, 0,  3.00000)    # Writes 3
    worksheet.write(4, 0,  3.00001)    # Writes 3.00001
    worksheet.write(5, 0,  3.14159)    # TeX revision no.?
    
    
    # Write some formulas
    worksheet.write(7, 0,  '=A3 + A6')
    worksheet.write(8, 0,  '=IF(A5>3,"Yes", "No")')
    
    
    # Write a hyperlink
    worksheet.write(10, 0, 'http://www.perl.com/')
    
    # File save
    workbook.close
    
    compare_file("perl_output/#{xlsfile}", xlsfile)
  end



  def compare_file(expected, target)
    fh_e = File.open(expected, "r")
    fh_t = File.open(target, "r")
    while true do
      e1 = fh_e.read(1)
      t1 = fh_t.read(1)
      if e1.nil?
        assert( t1.nil?, "#{expexted} is EOF but #{target} is NOT EOF.")
        break
      elsif t1.nil?
        assert( e1.nil?, '#{target} is EOF but #{expected} is NOT EOF.')
        break
      end
      assert_equal(e1, t1, sprintf(" #{expected} = '%s' but #{target} = '%s'", e1, t1))
      break
    end
    fh_e.close
    fh_t.close
  end


end
