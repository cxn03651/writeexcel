#####################################################
# t_workbook.rb
#
# Test suite for the Workbook class (workbook.rb)
# Requires testunit 0.1.8 or greater to run properly
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

class TC_Workbook < Test::Unit::TestCase

   def setup
      @wb = Workbook.new("test.xls")
   end

   def test_add_worksheet
      assert_nothing_raised{ ws = @wb.add_worksheet }
   end

   def test_calc_sheet_offsets
      ws = @wb.add_worksheet
      assert_nothing_raised{ @wb.calc_sheet_offsets }
   end

   def test_store_window1
      assert_nothing_raised{ @wb.store_window1 }
   end

   def test_store_all_fonts
      assert_nothing_raised{ @wb.store_all_fonts }
   end

   def test_store_xf
      assert_nothing_raised{ @wb.store_xf(0xFFF5) }
   end

   def test_store_all_xfs
      assert_nothing_raised{ @wb.store_all_xfs }
   end

   def test_store_style
      assert_nothing_raised{ @wb.store_style }
   end

   def test_store_boundsheet
      assert_nothing_raised{ @wb.store_boundsheet("test",0) }
   end

   def test_add_format
      assert_nothing_raised{ @wb.add_format }
      assert_equal(2,@wb.formats.length,"Bad number of formats")
      assert_nothing_raised{ 
        @wb.add_format(:bold => true, :size => 10, :color => 'black',
                       :fg_color => 43, :align => 'top', :text_wrap => true,
                       :border => 1)
      }
      assert_equal(3,@wb.formats.length,"Bad number of formats")
   end

   def teardown
      @wb.close
      @wb = nil
      File.delete("test.xls") if File.exists?("test.xls")
   end

end
