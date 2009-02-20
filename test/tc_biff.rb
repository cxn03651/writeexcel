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
   end

   def test_store_bof_length
      assert_nothing_raised{ @biff.store_bof }
      assert_equal(12, @biff.datasize, "Bad data size after store_bof call")
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
      assert_equal(26, @biff.datasize, "Bad data size for mixed data")
   end

   def teardown
      @biff = nil
   end
end
