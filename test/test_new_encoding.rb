# -*- coding: utf-8 -*-
require 'helper'
require 'nkf'

class TC_new_encoding < Test::Unit::TestCase
  def setup
    ruby_18 do
      @kcode = $KCODE
      $KCODE = 'u'
    end
  end

  def teardown
    ruby_18 { $KCODE = @kcode }
  end

  def test_finder
    if RUBY_VERSION < '1.9'
      %w{ UTF_8 UTF-8 utf-8 utf }.each do |name|
        e = Encoding.find(name)
        assert_equal 'UTF_8', e.name
        assert_equal 2, e.value
      end
    end
  end

  def test_comparison
    if RUBY_VERSION < '1.9'
      assert_equal Encoding.find('US-ASCII'), Encoding.find('ASCII')
      assert_equal Encoding.find('ASCII-8BIT'), Encoding.find('BINARY')
      assert_not_equal Encoding.find('ASCII'), Encoding.find('UTF-8')
    end
  end
  
  
  # TEST WITH ALTERNATE Encoding NAMES
  def test_ascii_str_ascii_to_ascii
    str = ascii_str_enc_us_ascii
    assert_equal(Encoding::ASCII, str.encode('ASCII').encoding)
  end
  def test_ascii_str_ascii_to_us_ascii
    str = ascii_str_enc_us_ascii
    assert_equal(Encoding::US_ASCII, str.encode('US-ASCII').encoding)
  end

  def test_ascii_str_ascii_to_binary
    str = ascii_str_enc_us_ascii
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end
  def test_ascii_str_ascii_to_ascii_8bit
    str = ascii_str_enc_us_ascii
    assert_equal(Encoding::ASCII_8BIT, str.encode('ASCII-8BIT').encoding)
  end

  def test_ascii_str_utf8_to_ascii
    str = ascii_str_enc_utf8
    assert_equal(Encoding::ASCII, str.encode('US-ASCII').encoding)
  end

  def test_non_ascii_str_utf8_to_ascii
    str = non_ascii_str_enc_utf8
    assert_raise(Encoding::UndefinedConversionError) { str.encode('US-ASCII') }
  end

  def test_ascii_str_utf8_to_binary
    str = ascii_str_enc_utf8
    assert_equal(Encoding::BINARY, str.encode('ASCII-8BIT').encoding)
  end

  def test_non_ascii_str_utf8_to_binary
    str = non_ascii_str_enc_utf8
    assert_raise(Encoding::UndefinedConversionError) { str.encode('ASCII-8BIT') }
  end
  
  
  
  def ascii_str_enc_ascii
    str = ["abc"].pack('a*').encode('ASCII')
    str.encode('ASCII') if RUBY_VERSION < "1.9"
    str
  end

  def test_ascii_str_enc_ascii_is_ascii_encoding
    assert_equal(Encoding::ASCII, ascii_str_enc_ascii.encoding)
  end

  def ascii_str_enc_us_ascii
    str = ["abc"].pack('a*').encode('US-ASCII')
    str.encode('US-ASCII') if RUBY_VERSION < "1.9"
    str
  end

  def test_ascii_str_enc_ascii_is_ascii_encoding
    assert_equal(Encoding::ASCII, ascii_str_enc_us_ascii.encoding)
  end

  def ascii_str_enc_binary
    str = ["abc"].pack('a*').encode('BINARY')
    str.encode('BINARY') if RUBY_VERSION < "1.9"
    str
  end

  def test_ascii_str_enc_binary_is_binary_encoding
    assert_equal(Encoding::BINARY, ascii_str_enc_binary.encoding)
  end

  def non_ascii_str_enc_binary
    str = [0x80].pack('v*')
    str.force_encoding('BINARY')
    str
  end

  def test_non_ascii_str_enc_binary
    assert_equal(Encoding::BINARY, non_ascii_str_enc_binary.encoding)
  end

  def ascii_str_enc_utf8
    "abc"
  end

  def test_ascii_str_enc_utf8_is_utf8_encoding
    assert_equal(Encoding::UTF_8, ascii_str_enc_utf8.encoding)
  end

  def non_ascii_str_enc_utf8
    'あいう'
  end

  def test_non_ascii_str_enc_utf8
    assert_equal(Encoding::UTF_8, non_ascii_str_enc_utf8.encoding)
  end

  def ascii_str_enc_eucjp
    str = "abc".encode('EUCJP')
    str
  end

  def test_ascii_str_enc_eucjp_is_eucjp_encoding
    assert_equal(Encoding::EUCJP, ascii_str_enc_eucjp.encoding)
  end

  def non_ascii_str_enc_eucjp
    str = 'あいう'.encode('EUCJP')
    str
  end

  def test_non_ascii_str_enc_eucjp
    assert_equal(Encoding::EUCJP, non_ascii_str_enc_eucjp.encoding)
  end

  def ascii_str_enc_sjis
    str = "abc".encode('SJIS')
    str
  end

  def test_ascii_str_enc_sjis_is_sjis_encoding
    assert_equal(Encoding::SJIS, ascii_str_enc_sjis.encoding)
  end

  def non_ascii_str_enc_sjis
    str = 'あいう'.encode('SJIS')
    str
  end

  def test_non_ascii_str_enc_sjis
    assert_equal(Encoding::SJIS, non_ascii_str_enc_sjis.encoding)
  end

  def ascii_str_enc_utf16le
    str = NKF.nkf('-w16L0 -m0 -W', "abc")
    str.force_encoding('UTF_16LE') if RUBY_VERSION < "1.9"
    str
  end

  def test_ascii_str_enc_utf16le_is_utf16le_encoding
    assert_equal(Encoding::UTF_16LE, ascii_str_enc_utf16le.encoding)
  end

  def non_ascii_str_enc_utf16le
    str = NKF.nkf('-w16L0 -m0 -W', 'あいう')
    str.force_encoding('UTF_16LE') if RUBY_VERSION < "1.9"
    str
  end

  def test_non_ascii_str_enc_utf16le
    assert_equal(Encoding::UTF_16LE, non_ascii_str_enc_utf16le.encoding)
  end

  def ascii_str_enc_utf16be
    str = NKF.nkf('-w16B0 -m0 -W', "abc")
    str.force_encoding('UTF_16BE') if RUBY_VERSION < "1.9"
    str
  end

  def test_ascii_str_enc_utf16be_is_utf16be_encoding
    assert_equal(Encoding::UTF_16BE, ascii_str_enc_utf16be.encoding)
  end

  def non_ascii_str_enc_utf16be
    str = NKF.nkf('-w16B0 -m0 -W', 'あいう')
    str.force_encoding('UTF_16BE') if RUBY_VERSION < "1.9"
    str
  end
  
end
