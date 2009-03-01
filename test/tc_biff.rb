#######################################################
# tc_biff.rb
#
# Test suite for the BIFFWriter class (biffwriter.rb)
  #######################################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i 
   Dir.chdir("..") if base == "test"
   $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
   Dir.chdir("test") rescue nil
end

require "test/unit"
require "biffwriter"

class TC_BIFFWriter < Test::Unit::TestCase
   def setup
      @biff = BIFFWriter.new
      @ruby_file = "delete_me"
   end

   def test_append_no_error
      assert_nothing_raised{ @biff.append("World") }
   end

   def test_prepend_no_error
      assert_nothing_raised{ @biff.prepend("Hello") }
   end

   def test_data_added
      assert_nothing_raised{ @biff.append("Hello", "World") }
      assert_equal("HelloWorld", @biff.data, "Bad data contents")
      assert_equal(10, @biff.datasize, "Bad data size")
   end

   def test_data_prepended

      assert_nothing_raised{ @biff.append("Hello") }
      assert_nothing_raised{ @biff.prepend("World") }
      assert_equal("WorldHello", @biff.data, "Bad data contents")
      assert_equal(10, @biff.datasize, "Bad data size")
   end

   def test_store_bof_length
      assert_nothing_raised{ @biff.store_bof }
      assert_equal(20, @biff.datasize, "Bad data size after store_bof call")
   end

   def test_store_eof_length
      assert_nothing_raised{ @biff.store_eof }
      assert_equal(4, @biff.datasize, "Bad data size after store_eof call")
   end

   def test_datasize_mixed
      assert_nothing_raised{ @biff.append("Hello") }
      assert_nothing_raised{ @biff.prepend("World") }
      assert_nothing_raised{ @biff.store_bof }
      assert_nothing_raised{ @biff.store_eof }
      assert_equal(34, @biff.datasize, "Bad data size for mixed data")
   end

   def test_add_continue
      perl_file = "perl_output/biff_add_continue_testdata"
      size = File.size(perl_file)
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@biff.add_continue('testdata'))
      @fh.close
      rsize = File.size(@ruby_file)
      assert_equal(size,rsize,"File sizes not the same")
   end

   def teardown
      @biff = nil
#      File.delete(@ruby_file) if File.exist?(@ruby_file)
   end
end
