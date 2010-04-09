# -*- coding: utf-8 -*-
require 'helper'
require 'nkf'

class TC_Compatibility < Test::Unit::TestCase
  def setup
    @kcode = $KCODE
    $KCODE = 'u'
  end

  def teardown
    $KCODE = @kcode
  end

  def test_decode_to_utf16le
    str = 'あいう'
    utf16le = NKF.nkf('-w16L0 -m0 -W', str)
    assert_equal(utf16le, str.encode('UTF-16LE'))
    assert_equal(utf16le, str.encode('utf-16le'))
  end

  def test_decode_to_utf16be
    str = 'あいう'
    utf16le = NKF.nkf('-w16B0 -m0 -W', str)
    assert_equal(utf16le, str.encode('UTF-16BE'))
    assert_equal(utf16le, str.encode('utf-16be'))
  end

  def test_encoding
    str     = 'あいう'
    utf8    = str
#    utf16le = NKF.nkf('-w16L0 -m0 -W', str)
#    utf16be = NKF.nkf('-w16B0 -m0 -W', str)
    assert_equal(Encoding::UTF_8, utf8.encoding)
  end

  def test_ord
    a = 'a'
    abc = 'abc'
    assert_equal(97, a.ord, "#{a}.ord faild\n")
    assert_equal(97, abc.ord, "#{abc}.ord faild\n")
  end

  def test_force_encodig
    str = 'あいう'
    org_dump = unpack_record(str)
    asc8_dump = unpack_record(str.force_encoding('ASCII-8BIT'))
    assert_equal(org_dump, asc8_dump)
  end
end
