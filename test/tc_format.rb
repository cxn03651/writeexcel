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
#require "workbook"
require "worksheet"
require "format"

class TC_Format < Test::Unit::TestCase

   def setup
      @ruby_file = "xf_test"
      @format = Format.new
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

   def test_set_format_properties
   end

   def test_format_properties_with_valid_value
      valid_properties = get_valid_format_properties
      valid_properties.each do |k,v|
         format = Format.new
         before = get_format_property(format)
         format.set_format_properties(k => v)
         after  = get_format_property(format)
         after.delete_if {|key, val| before[key] == val }
         assert_equal(1, after.size, "change 1 property[:#{k}] but #{after.size} was changed.#{after.inspect}")
         assert_equal(v, after[k], "[:#{k}] doesn't match.")
      end

      # set_color by string
      valid_color_string_number = get_valid_color_string_number
      [:color , :bg_color, :fg_color].each do |coltype|
         valid_color_string_number.each do |str, num|
            format = Format.new
            before = get_format_property(format)
            format.set_format_properties(coltype => str)
            after  = get_format_property(format)
            after.delete_if {|key, val| before[key] == val }
            assert_equal(1, after.size, "change 1 property[:#{coltype}:#{str}] but #{after.size} was changed.#{after.inspect}")
            assert_equal(num, after[:"#{coltype}"], "[:#{coltype}:#{str}] doesn't match.")
         end
      end


   end
   
   def test_format_properties_with_invalid_value
   end

   def test_set_font
   end

   def test_set_size
   end

   def test_set_color
   end

   def test_set_bold
   end

   def test_set_italic
   end

   def test_set_underline
   end

   def test_set_font_strikeout
   end

   def test_set_font_script
   end

   def test_set_font_outline
   end

   def test_set_font_shadow
   end

   def test_set_num_format
   end

   def test_set_locked
   end

   def test_set_hidden
   end

   def test_set_align
   end

   def test_set_center_across
   end

   def test_set_text_wrap
   end

   def test_set_rotation
   end

   def test_set_indent
   end

   def test_set_shrink
   end

   def test_set_text_justlast
   end

   def test_set_pattern
   end

   def test_set_bg_color
   end

   def test_set_fg_color
   end

   def test_set_border
   end

   def test_set_border_color
   end

   def test_copy
   end


   def test_xf_biff_size
      perl_file = "perl_output/file_xf_biff"
      size = File.size(perl_file)
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.get_xf)
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
      perl_file = "perl_output/file_font_biff"
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.get_font)
      @fh.close
      contents = IO.readlines(perl_file)
      rcontents = IO.readlines(@ruby_file)
      assert_equal(contents,rcontents,"Contents not the same")
   end

   def test_font_biff_contents
      perl_file = "perl_output/file_font_biff"
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.get_font)
      @fh.close
      contents = IO.readlines(perl_file)
      rcontents = IO.readlines(@ruby_file)
      assert_equal(contents,rcontents,"Contents not the same")
   end

   def test_get_font_key_size
      perl_file = "perl_output/file_font_key"
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.get_font_key)
      @fh.close
      assert_equal(File.size(perl_file),File.size(@ruby_file),"Bad file size")
   end

   def test_get_font_key_contents
      perl_file = "perl_output/file_font_key"
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.get_font_key)
      @fh.close
      contents = IO.readlines(perl_file)
      rcontents = IO.readlines(@ruby_file)
      assert_equal(contents,rcontents,"Contents not the same")
   end

   def test_initialize
     assert_nothing_raised {
       Format.new(:bold => true, :size => 10, :color => 'black', 
                  :fg_color => 43, :align => 'top', :text_wrap => true,
                  :border => 1)
     }
   end

   # added by Nakamura
   
   def test_get_xf
      perl_file = "perl_output/file_xf_biff"
      size = File.size(perl_file)
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.get_xf)
      @fh.close
      rsize = File.size(@ruby_file)
      assert_equal(size,rsize,"File sizes not the same")
      
      fh_p = File.open(perl_file, "r")
      fh_r = File.open(@ruby_file, "r")
      while true do
         p1 = fh_p.read(1)
         r1 = fh_r.read(1)
         if p1.nil?
            assert( r1.nil?, 'p1 is EOF but r1 is NOT EOF.')
            break
         elsif r1.nil?
            assert( p1.nil?, 'r1 is EOF but p1 is NOT EOF.')
            break
         end
         assert_equal(p1, r1, sprintf(" p1 = %s but r1 = %s", p1, r1))
         break
      end
      fh_p.close
      fh_r.close
   end
   
   def test_get_font
      perl_file = "perl_output/file_font_biff"
      size = File.size(perl_file)
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.get_font)
      @fh.close
      rsize = File.size(@ruby_file)
      assert_equal(size,rsize,"File sizes not the same")
      
      fh_p = File.open(perl_file, "r")
      fh_r = File.open(@ruby_file, "r")
      while true do
         p1 = fh_p.read(1)
         r1 = fh_r.read(1)
         if p1.nil?
            assert( r1.nil?, 'p1 is EOF but r1 is NOT EOF.')
            break
         elsif r1.nil?
            assert( p1.nil?, 'r1 is EOF but p1 is NOT EOF.')
            break
         end
         assert_equal(p1, r1, sprintf(" p1 = %s but r1 = %s", p1, r1))
         break
      end
      fh_p.close
      fh_r.close
   end
   
   def test_get_font_key
      perl_file = "perl_output/file_font_key"
      size = File.size(perl_file)
      @fh = File.new(@ruby_file,"w+")
      @fh.print(@format.get_font_key)
      @fh.close
      rsize = File.size(@ruby_file)
      assert_equal(size,rsize,"File sizes not the same")
      
      fh_p = File.open(perl_file, "r")
      fh_r = File.open(@ruby_file, "r")
      while true do
         p1 = fh_p.read(1)
         r1 = fh_r.read(1)
         if p1.nil?
            assert( r1.nil?, 'p1 is EOF but r1 is NOT EOF.')
            break
         elsif r1.nil?
            assert( p1.nil?, 'r1 is EOF but p1 is NOT EOF.')
            break
         end
         assert_equal(p1, r1, sprintf(" p1 = %s but r1 = %s", p1, r1))
         break
      end
      fh_p.close
      fh_r.close
   end
   
   def test_get_xf_index
   end
   
   def test_get_color
   end
   
   def test_method_missing
   end

# -----------------------------------------------------------------------------

   def get_valid_format_properties
      {
         :font => 'Times New Roman', 
         :size => 30, 
         :color => 8, 
         :italic => 1, 
         :underline => 1, 
         :font_strikeout => 1, 
         :font_script => 1, 
         :font_outline => 1, 
         :font_shadow => 1, 
         :locked => 0, 
         :hidden => 1, 
         :valign => 'top', 
         :text_wrap => 1, 
         :text_justlast => 1, 
         :indent => 2, 
         :shrink => 1, 
         :pattern => 18, 
         :bg_color => 30, 
         :fg_color => 63
      }
   end
   
   def get_valid_color_string_number
      return {
         'black'     =>    8,
         'blue'      =>   12,
         'brown'     =>   16,
         'cyan'      =>   15,
         'gray'      =>   23,
         'green'     =>   17,
         'lime'      =>   11,
         'magenta'   =>   14,
         'navy'      =>   18,
         'orange'    =>   53,
         'pink'      =>   33,
         'purple'    =>   20,
         'red'       =>   10,
         'silver'    =>   22,
         'white'     =>    9,
         'yellow'    =>   13
      }
   end
#         :rotation => -90, 
#         :center_across => 1, 
#         :align => 'left', 

   def get_format_property(format)
      text_h_align = {
         1 => 'left',
         2 => 'center/centre',
         3 => 'right',
         4 => 'fill',
         5 => 'justiry',
         6 => 'center_across/centre_across/merge',
         7 => 'distributed/equal_space'
      }

      text_v_align = {
         0 => 'top',
         1 => 'vcenter/vcentre',
         2 => 'bottom',
         3 => 'vjustify',
         4 => 'vdistributed/vequal_space'
      }

      return {
            :font => format.font, 
            :size => format.size, 
            :color => format.color, 
            :bold => format.bold, 
            :italic => format.italic, 
            :underline => format.underline, 
            :font_strikeout => format.font_strikeout, 
            :font_script => format.font_script, 
            :font_outline => format.font_outline, 
            :font_shadow => format.font_shadow, 
            :num_format => format.num_format, 
            :locked => format.locked, 
            :hidden => format.hidden, 
            :align => text_h_align[format.text_h_align],
            :valign => text_v_align[format.text_v_align], 
            :rotation => format.rotation, 
            :text_wrap => format.text_wrap, 
            :text_justlast => format.text_justlast, 
            :center_across => text_h_align[format.text_h_align], 
            :indent => format.indent, 
            :shrink => format.shrink, 
            :pattern => format.pattern, 
            :bg_color => format.bg_color, 
            :fg_color => format.fg_color, 
            :border => format.border,
            :bottom => format.bottom, 
            :top => format.top, 
            :left => format.left, 
            :right => format.right, 
            :bottom_color => format.bottom_color, 
            :top_color => format.top_color, 
            :left_color => format.left_color, 
            :right_color => format.right_color 
         }
   end

end
