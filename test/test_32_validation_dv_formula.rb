##########################################################################
# test_32_validation_dv_formula.rb
#
# Tests for the Excel DVAL structure used in data validation.
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


class TC_validation_dv_formula < Test::Unit::TestCase

  def setup
    @test_file  = 'temp_test_file.xls'
    @workbook   = Excel.new(@test_file)
    @worksheet  = @workbook.add_worksheet
    @worksheet2 = @workbook.add_worksheet
  end

  def test_integer_values
    formula      = '10'

    caption    = " \tData validation: _pack_dv_formula('#{formula}')"
    bytes      = %w(
                    03 00 00 E0 1E 0A 00
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_decimal_values
    formula      = '1.2345'

    caption    = " \tData validation: _pack_dv_formula('#{formula}')"
    bytes      = %w(
                    09 00 E0 3F 1F 8D 97 6E 12 83 C0 F3 3F
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_date_values
    formula      = @worksheet.convert_date_time('2008-07-24T')

    caption    = " \tData validation: _pack_dv_formula('2008-07-24T')"
    bytes      = %w(
                    03 00 E0 3F 1E E5 9A
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_time_values
    formula      = @worksheet.convert_date_time('T12:00')

    caption    = " \tData validation: _pack_dv_formula('T12:00')"
    bytes      = %w(
                    09 00 E0 3F 1F 00 00 00 00 00 00 E0 3F
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_cell_reference_value_C9
    formula      = '=C9'

    caption    = " \tData validation: _pack_dv_formula('#{formula}')"
    bytes      = %w(
                    05 00 E0 3F 44 08 00 02 C0
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_cell_reference_value_E3_E6
    formula      = '=E3:E6'

    caption    = " \tData validation: _pack_dv_formula('#{formula}')"
    bytes      = %w(
                    09 00 0C 00 25 02 00 05 00 04 C0 04 C0
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_cell_reference_value_E3_E6_absolute
    formula      = '=$E$3:$E$6'

    caption    = " \tData validation: _pack_dv_formula('#{formula}')"
    bytes      = %w(
                    09 00 0C 00 25 02 00 05 00 04 00 04 00
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_list_values
    formula      = ['a', 'bb', 'ccc']

    caption    = " \tData validation: _pack_dv_formula(['a', 'bb', 'ccc'])"
    bytes      = %w(
                    0B 00 0C 00 17 08 00 61 00 62 62 00 63 63 63
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_empty_string
    formula      = ''

    caption    = " \tData validation: _pack_dv_formula('')"
    bytes      = %w(
                    00 00 00
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
    assert_equal(target, result, caption)
  end

  def test_undefined_value
    formula      = nil

    caption    = " \tData validation: _pack_dv_formula(nil)"
    bytes      = %w(
                    00 00 00
                 )

    # Zero out Excel's random unused word to allow comparison.
    bytes[2]   = '00'
    bytes[3]   = '00'
    target     = bytes.join(" ")

    result     = unpack_record(@worksheet.pack_dv_formula(formula))
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
