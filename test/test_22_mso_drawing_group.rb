##########################################################################
# test_22_mso_drawing_group.rb
#
# Tests for the internal methods used to write the MSODRAWINGGROUP record.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
#########################################################################
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


class TC_mso_drawing_group < Test::Unit::TestCase

  def setup
    @test_file  = 'temp_test_file.xls'
    @workbook   = Excel.new(@test_file)
    @worksheet1 = @workbook.add_worksheet
    @worksheet2 = @workbook.add_worksheet
    @worksheet3 = @workbook.add_worksheet
  end

  def test_01
    count = 1
    for i in 1 .. count
      @worksheet1.write_comment(i -1, 0, 'aaa')
    end
    @workbook.calc_mso_sizes

    caption = sprintf(" \tSheet1: %4d comments.", count)
    target  = %w(
        EB 00 5A 00 0F 00 00 F0 52 00 00 00 00 00 06 F0
        18 00 00 00 02 04 00 00 02 00 00 00 02 00 00 00
        01 00 00 00 01 00 00 00 02 00 00 00 33 00 0B F0
        12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
        C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
        00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
    ).join(' ')
    result = unpack_record(@workbook.add_mso_drawing_group)
    assert_equal(target, result, caption)


    # Test the parameters pass to the worksheets
    caption   = caption + ' (params)'
    result_ids = []
    target_ids = [
                1024, 1, 2, 1025,
                 ]

    @workbook.sheets.each do |sheet|
      sheet.object_ids.each {|id| result_ids.push(id) }
    end
    
    assert_equal(target_ids, result_ids, caption)

  end

  def test_02
    count     = 2
    for i in 1 .. count
      @worksheet1.write_comment(i -1, 0, 'aaa')
    end
    @workbook.calc_mso_sizes

    target  = %w(
        EB 00 5A 00 0F 00 00 F0 52 00 00 00 00 00 06 F0
        18 00 00 00 03 04 00 00 02 00 00 00 03 00 00 00
        01 00 00 00 01 00 00 00 03 00 00 00 33 00 0B F0
        12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
        C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
        00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
    ).join(' ')
    caption    = sprintf( " \tSheet1: %4d comments.", count)
    result     = unpack_record(@workbook.add_mso_drawing_group)
    assert_equal(target, result, caption)


    # Test the parameters pass to the worksheets
    caption   = caption + ' (params)'
    result_ids = []
    target_ids = [
                1024, 1, 3, 1026,
                 ]

    @workbook.sheets.each do |sheet|
      sheet.object_ids.each {|id| result_ids.push(id) }
    end
    assert_equal(target_ids, result_ids, caption)

  end



  ###############################################################################
  #
  # Unpack the binary data into a format suitable for printing in tests.
  #
  def unpack_record(data)
    data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
  end

end
