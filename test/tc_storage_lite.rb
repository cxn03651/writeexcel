$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require "test/unit"
require "writeexcel/storage_lite"

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

class TC_OLEStorageLite < Test::Unit::TestCase
  def setup
    @ole = OLEStorageLite.new
  end
  
  def teardown
  end
  
  def test_asc2ucs
    result = @ole.asc2ucs('Root Entry')
    target = %w(
        52 00 6F 00 6F 00 74 00 20 00 45 00 6E 00 74 00 72 00 79 00
      ).join(" ")
    assert_equal(target, unpack_record(result))
  end
  
  def test_ucs2asc
    strings = [
        'Root Entry',
        ''
      ]
    strings.each do |str|
      result = @ole.ucs2asc(@ole.asc2ucs(str))
      assert_equal(str, result)
    end
  end
end

class TC_OLEStorageLitePPSFile < Test::Unit::TestCase
  def setup
  end
  
  def teardown
  end

  def test_constructor
    data = [
        { :name => 'name', :data => 'data' },
        { :name => '',     :data => 'data' },
        { :name => 'name', :data => ''     },
        { :name => '',     :data => ''     },
      ]
    data.each do |d|
      olefile = OLEStorageLitePPSFile.new(d[:name])
      assert_equal(d[:name], olefile.name)
    end
    data.each do |d|
      olefile = OLEStorageLitePPSFile.new(d[:name], d[:data])
      assert_equal(d[:name], olefile.name)
      assert_equal(d[:data], olefile.data)
    end
  end

  def test_append_no_file
    olefile = OLEStorageLitePPSFile.new('name')
    assert_equal('', olefile.data)

    data = [ "data", "\r\n", "\r", "\n" ]
    data.each do |d|
      olefile = OLEStorageLitePPSFile.new('name')
      olefile.append(d)
      assert_equal(d, olefile.data)
    end
  end

  def test_append_tempfile
    data = [ "data", "\r\n", "\r", "\n" ]
    data.each do |d|
      olefile = OLEStorageLitePPSFile.new('name')
      olefile.set_file
      pps_file = olefile.pps_file
    
      olefile.append(d)
      pps_file.open
      pps_file.binmode
      assert_equal(d, pps_file.read)
    end
  end

  def test_append_stringio
    data = [ "data", "\r\n", "\r", "\n" ]
    data.each do |d|
      sio = StringIO.new
      olefile = OLEStorageLitePPSFile.new('name')
      olefile.set_file(sio)
      pps_file = olefile.pps_file
    
      olefile.append(d)
      pps_file.rewind
      assert_equal(d, pps_file.read)
    end
  end
end
