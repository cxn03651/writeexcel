######################################################################
# tc_worksheet.rb
#
# Test suite for the Worksheet class (worksheet.rb).
#
# I do lots of octal dump comparisons because I don't trust simply
# visually comparing the file contents (you never know if there's
# a hidden space or newline somewhere).  Besides, it can't hurt.
#######################################################################
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
require "ftools"

class TC_Worksheet < Test::Unit::TestCase
   def setup
      @ws      = Worksheet.new("test",0)
      @perldir = "perl_output/"
      @format  = Format.new(:color=>"green")
   end

   def test_methods_exist
      assert_respond_to(@ws, :write)
      assert_respond_to(@ws, :write_blank)
      assert_respond_to(@ws, :write_row)
      assert_respond_to(@ws, :write_column)
   end

   def test_format_rectangle
      assert_respond_to(@ws, :format_rectangle)
      assert_nothing_raised{ @ws.format_rectangle(0, 0, 2, 2, @format) }
      assert_raises(TypeError){ @ws.format_rectangle(0, 0, 2, 2, "bogus") }
   end

   def test_methods_no_error
      assert_nothing_raised{ @ws.write(0,0,nil) }
      assert_nothing_raised{ @ws.write(0,0,"Hello") }
      assert_nothing_raised{ @ws.write(0,0,888) }
      assert_nothing_raised{ @ws.write_row(0,0,nil) }
      assert_nothing_raised{ @ws.write_row(0,0,["one","two","three"]) }
      assert_nothing_raised{ @ws.write_row(0,0,[1,2,3]) }
      assert_nothing_raised{ @ws.write_column(0,0,nil) }
      assert_nothing_raised{ @ws.write_column(0,0,["one","two","three"]) }
      assert_nothing_raised{ @ws.write_column(0,0,[1,2,3]) }
      assert_nothing_raised{ @ws.write_blank(0,0,nil) }
      assert_nothing_raised{ @ws.write_url(0,0,"http://www.ruby-lang.org") }
   end

   def test_store_dimensions
      file = "delete_this"
      File.open(file,"w+"){ |f| f.print @ws.store_dimensions }
      pf = @perldir + "ws_store_dimensions"
      p_od = IO.readlines(pf).to_s.dump
      r_od = IO.readlines(file).to_s.dump
      assert_equal(2,File.size(file),"Bad file size")
      assert_equal(p_od, r_od,"Octal dumps are not identical")
      File.delete(file)
   end

   def test_store_window2
      file = "delete_this"
      File.open(file,"w+"){ |f| f.print @ws.store_window2 }
      pf = @perldir + "ws_store_window2"
      p_od = IO.readlines(pf).to_s.dump
      r_od = IO.readlines(file).to_s.dump
      assert_equal(2,File.size(file),"Bad file size")
      assert_equal(p_od, r_od,"Octal dumps are not identical")
      File.delete(file)
   end

   def test_store_selection
      file = "delete_this"
      File.open(file,"w+"){ |f| f.print @ws.store_selection }
      pf = @perldir + "ws_store_selection"
      p_od = IO.readlines(pf).to_s.dump
      r_od = IO.readlines(file).to_s.dump
      assert_equal(2,File.size(file),"Bad file size")
      assert_equal(p_od, r_od,"Octal dumps are not identical")
      File.delete(file)
   end

   def test_store_colinfo_output
      file = "delete_this"
      File.open(file,"w+"){ |f| f.print @ws.store_colinfo }
      pf = @perldir + "ws_colinfo"
      p_od = IO.readlines(pf).to_s.dump
      r_od = IO.readlines(file).to_s.dump
      assert_equal(2,File.size(file),"Invalid size for store_colinfo")
      assert_equal(p_od,r_od,"Perl and Ruby octal dumps don't match")
      File.delete(file)
   end

   def test_write_syntax
      assert_nothing_raised{@ws.write(0,0,"Hello")}
      assert_nothing_raised{@ws.write(0,0,666)}
   end

   def test_write_string
      file = "delete_this"
      @ws.write(0,0,"Hello")
      File.open(file,"w+"){ |f| f.print(@ws.data) }
      assert_equal(17,File.size(file),"incorrect file size")
      File.delete(file)
   end

   def test_write_number
      file = "delete_this"
      @ws.write(0,0,6789)
      File.open(file,"w+"){ |f| f.print(@ws.data) }
      assert_equal(18,File.size(file),"incorrect file size")
      File.delete(file)
   end

   def test_format_row
      assert_nothing_raised{ @ws.format_row(3,nil) }
      assert_nothing_raised{ @ws.format_row(2,10) }
      assert_nothing_raised{ @ws.format_row(1,25,@format) }
      assert_nothing_raised{ @ws.format_row(4..6,nil,@format) }
   end

   def test_format_column
      assert_nothing_raised{ @ws.format_column(3,nil) }
      assert_nothing_raised{ @ws.format_column(2,10) }
      assert_nothing_raised{ @ws.format_column(1,25,@format) }
      assert_nothing_raised{ @ws.format_column(4..6,nil,@format) }
   end

   def teardown
      @ws     = nil
      @format = nil
   end
end
