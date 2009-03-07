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
require "FileUtils"

class TC_Worksheet < Test::Unit::TestCase
  def setup
    @sheetname = 'test'
    @ws      = Worksheet.new(@sheetname,0)
    @perldir = "perl_output/"
    @format  = Format.new(:color=>"green")
  end

  def teardown
    @ws     = nil
    @format = nil
  end

  def test_methods_exist
    assert_respond_to(@ws, :write)
    assert_respond_to(@ws, :write_blank)
    assert_respond_to(@ws, :write_row)
    assert_respond_to(@ws, :write_col)
  end

  def test_methods_no_error
    assert_nothing_raised{ @ws.write(0,0,nil) }
    assert_nothing_raised{ @ws.write(0,0,"Hello") }
    assert_nothing_raised{ @ws.write(0,0,888) }
    assert_nothing_raised{ @ws.write_row(0,0,["one","two","three"]) }
    assert_nothing_raised{ @ws.write_row(0,0,[1,2,3]) }
    assert_nothing_raised{ @ws.write_col(0,0,["one","two","three"]) }
    assert_nothing_raised{ @ws.write_col(0,0,[1,2,3]) }
    assert_nothing_raised{ @ws.write_blank(0,0,nil) }
    assert_nothing_raised{ @ws.write_url(0,0,"http://www.ruby-lang.org") }
  end

  def test_write_syntax
    assert_nothing_raised{@ws.write(0,0,"Hello")}
    assert_nothing_raised{@ws.write(0,0,666)}
  end

  def test_store_dimensions
    file = "delete_this"
    File.open(file,"w+"){ |f| f.print @ws.store_dimensions }
    pf = @perldir + "ws_store_dimensions"
    p_od = IO.readlines(pf).to_s.dump
    r_od = IO.readlines(file).to_s.dump
    assert_equal_filesize(pf ,file, "Invalid size for store_selection")
    assert_equal(p_od, r_od,"Octal dumps are not identical")
    File.delete(file)
  end

  def test_store_colinfo
    file = "delete_this"
    File.open(file,"w+"){ |f| f.print @ws.store_colinfo }
    pf = @perldir + "ws_store_colinfo"
    p_od = IO.readlines(pf).to_s.dump
    r_od = IO.readlines(file).to_s.dump
    assert_equal_filesize(pf, file, "Invalid size for store_colinfo")
    assert_equal(p_od,r_od,"Perl and Ruby octal dumps don't match")
    File.delete(file)
  end

  def test_store_selection
    file = "delete_this"
    File.open(file,"w+"){ |f| f.print @ws.store_selection(1,1,2,2) }
    pf = @perldir + "ws_store_selection"
    p_od = IO.readlines(pf).to_s.dump
    r_od = IO.readlines(file).to_s.dump
    assert_equal_filesize(pf, file, "Invalid size for store_selection")
    assert_equal(p_od, r_od,"Octal dumps are not identical")
    File.delete(file)
  end

  def test_store_filtermode
    file = "delete_this"
    File.open(file,"w+"){ |f| f.print @ws.store_filtermode }
    pf = @perldir + "ws_store_filtermode_off"
    p_od = IO.readlines(pf).to_s.dump
    r_od = IO.readlines(file).to_s.dump
    assert_equal_filesize(pf, file, "Invalid size for store_filtermode_off")
    assert_equal(p_od, r_od,"Octal dumps are not identical")
    File.delete(file)

    @ws.autofilter(1,1,2,2)
    @ws.filter_column(1, 'x < 2000')
    File.open(file,"w+"){ |f| f.print @ws.store_filtermode }
    pf = @perldir + "ws_store_filtermode_on"
    p_od = IO.readlines(pf).to_s.dump
    r_od = IO.readlines(file).to_s.dump
    assert_equal_filesize(pf, file, "Invalid size for store_filtermode_off")
    assert_equal(p_od, r_od,"Octal dumps are not identical")
    File.delete(file)
  end

  def test_new
    assert_equal(@sheetname, @ws.name)
  end


  def assert_equal_filesize(target, test, msg = "Bad file size")
    assert_equal(File.size(target),File.size(test),msg)
  end

end
