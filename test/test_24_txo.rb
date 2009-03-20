##########################################################################
# test_24_txo.rb
#
# Tests for some of the internal method used to write the NOTE record that
# is used in cell comments.
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
require 'writeexcel'
include Spreadsheet


class TC_txo < Test::Unit::TestCase

  def setup
    t = Time.now.strftime("%Y%m%d")
    path = "temp#{t}-#{$$}-#{rand(0x100000000).to_s(36)}"
    @test_file           = File.join(Dir.tmpdir, path)
    @workbook   = WriteExcel.new(@test_file)
    @worksheet  = @workbook.add_worksheet
  end

  def teardown
    @workbook.close
    File.unlink(@test_file) if FileTest.exist?(@test_file)
  end

  def test_txo
    string     = 'aaa'
    caption    = " \t_store_txo()"
    target     = %w(
                    B6 01 12 00 12 02 00 00 00 00 00 00 00 00 03 00
                    10 00 00 00 00 00
                   ).join(' ')

    result     = unpack_record(@worksheet.store_txo(string.length))
    assert_equal(target, result, caption)
  end

  def test_first_continue_record_after_txo
    string     = 'aaa'
    caption    = " \t_store_txo_continue_1()"
    target     = %w(
                    3C 00 04 00 00 61 61 61
                   ).join(' ')

    result     = unpack_record(@worksheet.store_txo_continue_1(string))
    assert_equal(target, result, caption)
  end

  def test_second_continue_record_after_txo
    string     = 'aaa'
    caption    = " \t_store_txo_continue_2()"
    target     = %w(
                    3C 00 10 00 00 00 00 00 00 00 00 00 03 00 00 00
                    00 00 00 00
                   ).join(' ')
    formats    = [
                    [0,             0],
                    [string.length, 0]
                 ]

    result     = unpack_record(@worksheet.store_txo_continue_2(formats))
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
