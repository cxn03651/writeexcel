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

   def test_new
      assert_kind_of(Workbook, @wb)
   end

   def test_add_worksheet
      sheetnames = ['sheet1', 'sheet2']
      (0 .. sheetnames.size-1).each do |i|
         sheets = @wb.sheets
         assert_equal(i, sheets.size)
         @wb.add_worksheet(sheetnames[i])
         sheets = @wb.sheets
         assert_equal(i+1, sheets.size)
      end
   end

   def test_set_tempdir
      # after shees added, call set_tempdir raise RuntimeError
      wb1 = Workbook.new('wb1')
      wb1.add_worksheet('name')
      assert_raise(RuntimeError, "already sheet exists, but set_tempdir() doesn't raise"){
         wb1.set_tempdir
      }

      # invalid dir raise RuntimeError
      wb2 = Workbook.new('wb2')
      while true do
         dir = Time.now.to_s
         break unless FileTest.directory?(dir)
         sleep 0.1
      end
      assert_raise(RuntimeError, "#{dir} is not valid directory"){
         wb2.set_tempdir(dir)
      }
   end

=begin
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
=end
end
