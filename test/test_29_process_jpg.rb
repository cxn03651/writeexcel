##########################################################################
# test_29_process_jpg.rb
#
# Tests for the JPEG width and height processing.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
#########################################################################
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require "test/unit"
require 'writeexcel'

class TC_29_process_jpg < Test::Unit::TestCase

  def setup
    t = Time.now.strftime("%Y%m%d")
    path = "temp#{t}-#{$$}-#{rand(0x100000000).to_s(36)}"
    @test_file           = File.join(Dir.tmpdir, path)
    @workbook   = Spreadsheet::WriteExcel.new(@test_file)
    @type       = 5   # Excel Blip type (MSOBLIPTYPE).
  end

  def teardown
    @workbook.close
    File.unlink(@test_file) if FileTest.exist?(@test_file)
  end

  def test_valid_jpg_image_1
    testname = '3w x 5h jpeg image.'
    
    data = %w(
      FF D8 FF E0 00 10 4A 46 49 46 00 01 01 01 00 60
      00 60 00 00 FF DB 00 43 00 06 04 05 06 05 04 06
      06 05 06 07 07 06 08 0A 10 0A 0A 09 09 0A 14 0E
      0F 0C 10 17 14 18 18 17 14 16 16 1A 1D 25 1F 1A
      1B 23 1C 16 16 20 2C 20 23 26 27 29 2A 29 19 1F
      2D 30 2D 28 30 25 28 29 28 FF DB 00 43 01 07 07
      07 0A 08 0A 13 0A 0A 13 28 1A 16 1A 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 FF C0
      00 11 08 00 05 00 03 03 01 22 00 02 11 01 03 11
      01 FF C4 00 15 00 01 01 00 00 00 00 00 00 00 00
      00 00 00 00 00 00 00 07 FF C4 00 14 10 01 00 00
      00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF C4
      00 15 01 01 01 00 00 00 00 00 00 00 00 00 00 00
      00 00 00 06 08 FF C4 00 14 11 01 00 00 00 00 00
      00 00 00 00 00 00 00 00 00 00 00 FF DA 00 0C 03
      01 00 02 11 03 11 00 3F 00 9D 00 1C A4 5F FF D9
    )
    image = [data.join('')].pack('H*')

    expected = [@type, 3, 5]
    result   = @workbook.process_jpg(image, 'test.jpg')
    assert_equal(expected, result, " \t" + testname)
  end

  def test_valid_jpg_image_2
    testname = '5w x 3h jpeg image.'
    
    data = %w(
      FF D8 FF E0 00 10 4A 46 49 46 00 01 01 01 00 60
      00 60 00 00 FF DB 00 43 00 06 04 05 06 05 04 06
      06 05 06 07 07 06 08 0A 10 0A 0A 09 09 0A 14 0E
      0F 0C 10 17 14 18 18 17 14 16 16 1A 1D 25 1F 1A
      1B 23 1C 16 16 20 2C 20 23 26 27 29 2A 29 19 1F
      2D 30 2D 28 30 25 28 29 28 FF DB 00 43 01 07 07
      07 0A 08 0A 13 0A 0A 13 28 1A 16 1A 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 FF C0
      00 11 08 00 03 00 05 03 01 22 00 02 11 01 03 11
      01 FF C4 00 15 00 01 01 00 00 00 00 00 00 00 00
      00 00 00 00 00 00 00 07 FF C4 00 14 10 01 00 00
      00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF C4
      00 15 01 01 01 00 00 00 00 00 00 00 00 00 00 00
      00 00 00 06 08 FF C4 00 14 11 01 00 00 00 00 00
      00 00 00 00 00 00 00 00 00 00 00 FF DA 00 0C 03
      01 00 02 11 03 11 00 3F 00 9D 00 1C A4 5F FF D9
    )
    image = [data.join('')].pack('H*')

    expected = [@type, 5, 3]
    result   = @workbook.process_jpg(image, 'test.jpg')
    assert_equal(expected, result, " \t" + testname)
  end

  def test_valid_jpg_image_3_ffco_marker_missing
    testname = 'FFCO marker missing in image.'
    
    data = %w(
      FF D8 FF E0 00 10 4A 46 49 46 00 01 01 01 00 60
      00 60 00 00 FF DB 00 43 00 06 04 05 06 05 04 06
      06 05 06 07 07 06 08 0A 10 0A 0A 09 09 0A 14 0E
      0F 0C 10 17 14 18 18 17 14 16 16 1A 1D 25 1F 1A
      1B 23 1C 16 16 20 2C 20 23 26 27 29 2A 29 19 1F
      2D 30 2D 28 30 25 28 29 28 FF DB 00 43 01 07 07
      07 0A 08 0A 13 0A 0A 13 28 1A 16 1A 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 28 28
      28 28 28 28 28 28 28 28 28 28 28 28 28 28 FF C1
      00 11 08 00 03 00 05 03 01 22 00 02 11 01 03 11
      01 FF C4 00 15 00 01 01 00 00 00 00 00 00 00 00
      00 00 00 00 00 00 00 07 FF C4 00 14 10 01 00 00
      00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF C4
      00 15 01 01 01 00 00 00 00 00 00 00 00 00 00 00
      00 00 00 06 08 FF C4 00 14 11 01 00 00 00 00 00
      00 00 00 00 00 00 00 00 00 00 00 FF DA 00 0C 03
      01 00 02 11 03 11 00 3F 00 9D 00 1C A4 5F FF D9
    )
    image = [data.join('')].pack('H*')

    assert_raise(RuntimeError, " \t" + testname) {
      @workbook.process_jpg(image, 'test.jpg')
    }
  end

  def test_invalid_jpeg_image
    testname = 'empty image'
    image    = ''
    
    assert_raise(RuntimeError, " \t" + testname) {
      @workbook.process_jpg(image, 'test.jpg')
    }
  end

end
