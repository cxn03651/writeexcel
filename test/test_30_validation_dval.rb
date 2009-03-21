##########################################################################
# test_30_validation_dval.rb
#
# Tests for the Excel DVAL structure used in data validation.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
#########################################################################
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require "test/unit"
require 'writeexcel'

class TC_validation_dval < Test::Unit::TestCase

  def setup
    t = Time.now.strftime("%Y%m%d")
    path = "temp#{t}-#{$$}-#{rand(0x100000000).to_s(36)}"
    @test_file           = File.join(Dir.tmpdir, path)
    @workbook   = Spreadsheet::WriteExcel.new(@test_file)
    @worksheet  = @workbook.add_worksheet
  end

  def teardown
    @workbook.close
    File.unlink(@test_file) if FileTest.exist?(@test_file)
  end

  def test_1
    obj_id     = 1
    dv_count   = 1

    caption    = " \tData validation: _store_dval(#{obj_id}, #{dv_count})"
    target     = %w(
                   B2 01 12 00 04 00 00 00 00 00 00 00 00 00 01 00
                   00 00 01 00 00 00
                 ).join(' ')

    result     = unpack_record(@worksheet.store_dval(obj_id, dv_count))
    assert_equal(target, result, caption)
  end

  def test_2
    obj_id     = -1
    dv_count   = 1

    caption    = " \tData validation: _store_dval(#{obj_id}, #{dv_count})"
    target     = %w(
                   B2 01 12 00 04 00 00 00 00 00 00 00 00 00 FF FF
                   FF FF 01 00 00 00
                 ).join(' ')

    result     = unpack_record(@worksheet.store_dval(obj_id, dv_count))
    assert_equal(target, result, caption)
  end

  def test_3
    obj_id     = 1
    dv_count   = 2

    caption    = " \tData validation: _store_dval(#{obj_id}, #{dv_count})"
    target     = %w(
                   B2 01 12 00 04 00 00 00 00 00 00 00 00 00 01 00
                   00 00 02 00 00 00
                 ).join(' ')

    result     = unpack_record(@worksheet.store_dval(obj_id, dv_count))
    assert_equal(target, result, caption)
  end

  ###############################################################################
  #
  # Unpack the binary data into a format suitable for printing in tests.
  #
  def unpack_record(data)
    data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
  end

end
