#######################################################
# test_00_IEEE_double.rb
#
# Check if "pack" gives the required IEEE 64bit float
#######################################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i
  Dir.chdir("..") if base == "test"
  $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
  Dir.chdir("test") rescue nil
end

require "test/unit"

class TC_BIFFWriter < Test::Unit::TestCase

  def test_IEEE_double
    teststr = [1.2345].pack("d")
    hexdata = [0x8D, 0x97, 0x6E, 0x12, 0x83, 0xC0, 0xF3, 0x3F]
    number  = hexdata.pack("C8")

    assert(number == teststr || number == teststr.reverse, "Not Little/Big endian. Give up.")
  end
end
