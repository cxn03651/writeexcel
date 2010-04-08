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
end
