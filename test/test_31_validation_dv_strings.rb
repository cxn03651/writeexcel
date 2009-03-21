##########################################################################
# test_31_validation_dv_strings.rb
#
# Tests for the packed caption/message strings used in the Excel DV structure
# as part of data validation.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
#########################################################################
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require "test/unit"
require 'writeexcel'

class TC_validation_dv_strings < Test::Unit::TestCase

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

  def test_empty_string
    string      = ''
    max_length  = 32

    caption    = " \tData validation: _pack_dv_string('', #{max_length})"
    target     = %w(
                   01 00 00 00
                 ).join(' ')
    result     = unpack_record(@worksheet.pack_dv_string(string, max_length))
    assert_equal(target, result, caption)
  end

  def test_nil
    string      = nil
    max_length  = 32

    caption    = " \tData validation: _pack_dv_string('', #{max_length})"
    target     = %w(
                   01 00 00 00
                 ).join(' ')
    result     = unpack_record(@worksheet.pack_dv_string(string, max_length))
    assert_equal(target, result, caption)
  end

  def test_single_space
    string      = ' '
    max_length  = 32

    caption    = " \tData validation: _pack_dv_string('', #{max_length})"
    target     = %w(
                   01 00 00 20
                 ).join(' ')
    result     = unpack_record(@worksheet.pack_dv_string(string, max_length))
    assert_equal(target, result, caption)
  end

  def test_single_character
    string      = 'A'
    max_length  = 32

    caption    = " \tData validation: _pack_dv_string('', #{max_length})"
    target     = %w(
                   01 00 00 41
                 ).join(' ')
    result     = unpack_record(@worksheet.pack_dv_string(string, max_length))
    assert_equal(target, result, caption)
  end

  def test_string_longer_than_32_characters_for_dialog_captions
    string      = 'This string is longer than 32 characters'
    max_length  = 32

    caption    = " \tData validation: _pack_dv_string('', #{max_length})"
    target     = %w(
                   20 00 00 54 68 69 73 20
                   73 74 72 69 6E 67 20 69 73 20 6C 6F 6E 67 65 72
                   20 74 68 61 6E 20 33 32 20 63 68
                 ).join(' ')
    result     = unpack_record(@worksheet.pack_dv_string(string, max_length))
    assert_equal(target, result, caption)
  end

  def test_string_longer_than_32_characters_for_dialog_messages
    string      = 'ABCD' * 64
    max_length  = 255

    caption    = " \tData validation: _pack_dv_string('', #{max_length})"
    target     = %w(
                            FF 00 00 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43 44 41 42 43 44 41 42 43 44 41 42 43 44 41
                            42 43
                 ).join(' ')
    result     = unpack_record(@worksheet.pack_dv_string(string, max_length))
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
