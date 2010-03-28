$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require "test/unit"
require 'writeexcel'
require 'stringio'

  def unpack_record(data)
    data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
  end

###############################################################################
#
# A test for Chart.
#
# Tests for the Excel chart.rb methods.
#
# reverse(''), December 2009, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#
class TC_ChartGeneric < Test::Unit::TestCase
  def setup
    io = StringIO.new
    workbook = WriteExcel.new(io)
    @chart = Chart.new(workbook, '', 'chart', 0, 0, 0, 0)
  end

  ###############################################################################
  #
  # Test the _store_fbi method.
  #
  def test_store_fbi
    caption = " \tChart, _store_fbi()"
    expected = %w(
        60 10 0A 00 B8 38 A1 22 C8 00 00 00 05 00
      ).join(' ')
    got = unpack_record(@chart.store_fbi(5))
    assert_equal(expected, got, caption)

    # Try a different index.
    expected = %w(
        60 10 0A 00 B8 38 A1 22 C8 00 00 00 06 00
      ).join(' ')
    got = unpack_record(@chart.store_fbi(6))
    assert_equal(expected, got, caption)
  end

  ###############################################################################
  #
  # Test the _store_chart method.
  #
  def test_store_chart
    caption = " \tChart, _store_chart()";
    expected = %w(
        02 10 10 00 00 00 00 00 00 00 00 00 E0 51 DD 02
        38 B8 C2 01
      ).join(' ')
    got = unpack_record(@chart.store_chart)
    assert_equal(expected, got, caption)
 end
end
=begin
