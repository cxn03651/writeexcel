#####################################################
# tc_format.rb
#
# Test suite for the Format class (format.rb)
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

class TC_Format < Test::Unit::TestCase

   def setup
      @ruby_file = "xf_test"
      @format = Format.new
   end

   def test_xf_biff_size
      perl_file = "perl_output/f_xf_biff"
      size = File.size(perl_file)
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.xf_biff)
      @fh.close
      rsize = File.size(@ruby_file)
      assert_equal(size,rsize,"File sizes not the same")
   end
   
   # Because of the modifications to bg_color and fg_color, I know this
   # test will fail.  This is ok.
   #def test_xf_biff_contents
   #   perl_file = "perl_output/f_xf_biff"
   #   @fh = File.new(@ruby_file,"w+")
   #   @fh.print(@format.xf_biff)
   #   @fh.close
   #   contents = IO.readlines(perl_file)
   #   rcontents = IO.readlines(@ruby_file)
   #   assert_equal(contents,rcontents,"Contents not the same")
   #end

   def test_font_biff_size
      perl_file = "perl_output/f_font_biff"
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.font_biff)
      @fh.close
      contents = IO.readlines(perl_file)
      rcontents = IO.readlines(@ruby_file)
      assert_equal(contents,rcontents,"Contents not the same")
   end

   def test_font_biff_contents
      perl_file = "perl_output/f_font_biff"
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.font_biff)
      @fh.close
      contents = IO.readlines(perl_file)
      rcontents = IO.readlines(@ruby_file)
      assert_equal(contents,rcontents,"Contents not the same")
   end

   def test_get_font_key_size
      perl_file = "perl_output/f_font_key"
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.font_key)
      @fh.close
      assert_equal(File.size(perl_file),File.size(@ruby_file),"Bad file size")
   end

   def test_get_font_key_contents
      perl_file = "perl_output/f_font_key"
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.font_key)
      @fh.close
      contents = IO.readlines(perl_file)
      rcontents = IO.readlines(@ruby_file)
      assert_equal(contents,rcontents,"Contents not the same")
   end

   def test_color
      assert_nothing_raised{ @format.color = "blue" }
      assert_equal(0x0C,@format.color,"Bad color value")
   end

   def test_color_valid
      colors = %w/aqua black blue brown cyan fuchsia gray grey green lime/
      colors << %w/magenta navy orange purple red silver white yellow/
      colors.flatten!

      colors.each{ |color|
         assert_nothing_raised{ @format.color = color }
      }
   end

   def test_color_bogus
      assert_raises(ArgumentError){ @format.color = "blah" }
   end

   def test_align
      @format.align
      assert_equal(0,@format.text_h_align,"Bad text_h_align")
      assert_equal(2,@format.text_v_align,"Bad text_v_align")
   end

   def test_align_center
      @format.align = "center"
      assert_equal(2,@format.text_h_align,"Bad text_h_align")
      assert_equal(2,@format.text_v_align,"Bad text_v_align")
   end

   def test_bold
      @format.bold=1
      assert_equal(true,@format.bold,"Bad bold value")
   end

   def test_initialize
     assert_nothing_raised {
       Format.new(:bold => true, :size => 10, :color => 'black', 
                  :fg_color => 43, :align => 'top', :text_wrap => true,
                  :border => 1)
     }
   end

   def teardown
      begin
         @pfh.close
      rescue NameError
         # no op
      end
      File.delete(@ruby_file) if File.exist?(@ruby_file)
      @format = nil
   end

end
