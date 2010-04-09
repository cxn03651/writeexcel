# -*- coding: utf-8 -*-
require 'helper'
require 'nkf'

class TC_Compatibility < Test::Unit::TestCase
  def setup
    ruby_18 do
      @kcode = $KCODE
      $KCODE = 'u'
    end
  end

  def teardown
    ruby_18 { $KCODE = @kcode }
  end

  def test_encoding
    # String#encoding -> Encoding::***
    str     = 'あいう'
    utf8    = str
#    utf16le = NKF.nkf('-w16L0 -m0 -W', str)
#    utf16be = NKF.nkf('-w16B0 -m0 -W', str)
    assert_equal(Encoding::UTF_8, utf8.encoding)

    if RUBY_VERSION < "1.9"
      # String#encoding returns depending $KCODE
      # if string has not changed its encoding.
      str   = 'abc'
      utf8  = str
      assert_equal(Encoding::UTF_8, utf8.encoding)
    end
  end

  def ascii_str_enc_ascii
    str = ["abc"].pack('a*').encode('ASCII')
    str.encode('ASCII') if RUBY_VERSION < "1.9"
    str
  end

  def test_ascii_str_enc_ascii_is_ascii_encoding
    assert_equal(Encoding::ASCII, ascii_str_enc_ascii.encoding)
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

  def test_non_ascii_str_enc_utf16be
    assert_equal(Encoding::UTF_16BE, non_ascii_str_enc_utf16be.encoding)
  end

  def test_ascii_str_ascii_to_ascii
    str = ascii_str_enc_ascii
    assert_equal(Encoding::ASCII, str.encode('ASCII').encoding)
  end

  def test_ascii_str_ascii_to_binary
    str = ascii_str_enc_ascii
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end

  def test_ascii_str_ascii_to_utf8
    str = ascii_str_enc_ascii
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_8') }
  end

  def test_ascii_str_ascii_to_eucjp
    str = ascii_str_enc_ascii
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_ascii_str_ascii_to_sjis
    str = ascii_str_enc_ascii
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_ascii_str_ascii_to_utf16le
    str = ascii_str_enc_ascii
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16LE') }
  end

  def test_ascii_str_ascii_to_utf16be
    str = ascii_str_enc_ascii
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16BE') }
  end

  def test_ascii_str_binary_to_ascii
    str = ascii_str_enc_binary
    assert_equal(Encoding::ASCII, str.encode('ASCII').encoding)
  end

  def test_non_ascii_str_binary_to_ascii
    str = non_ascii_str_enc_binary
    assert_raise(Encoding::UndefinedConversionError){ str.encode('ASCII') }
  end

  def test_ascii_str_binary_to_binary
    str = ascii_str_enc_binary
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end

  def test_non_ascii_str_binary_to_binary
    str = non_ascii_str_enc_binary
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end

  def test_ascii_str_binary_to_utf8
    str = ascii_str_enc_binary
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_8') }
  end

  def test_non_ascii_str_binary_to_utf8
    str = non_ascii_str_enc_binary
    assert_raise(Encoding::ConverterNotFoundError,
                 Encoding::UndefinedConversionError) { str.encode('UTF_8') }
  end

  def test_ascii_str_binary_to_eucjp
    str = ascii_str_enc_binary
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_non_ascii_str_binary_to_eucjp
    str = non_ascii_str_enc_binary
    assert_raise(Encoding::ConverterNotFoundError,
                 Encoding::UndefinedConversionError) { str.encode('EUCJP') }
  end

  def test_ascii_str_binary_to_sjis
    str = ascii_str_enc_binary
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_non_ascii_str_binary_to_sjis
    str = non_ascii_str_enc_binary
    assert_raise(Encoding::ConverterNotFoundError,
                 Encoding::UndefinedConversionError) { str.encode('SJIS') }
  end

  def test_ascii_str_binary_to_utf16le
    str = ascii_str_enc_binary
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16LE') }
  end

  def test_non_ascii_str_binary_to_utf16le
    str = non_ascii_str_enc_binary
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16LE') }
  end

  def test_ascii_str_binary_to_utf16be
    str = ascii_str_enc_binary
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16BE') }
  end

  def test_non_ascii_str_binary_to_utf16be
    str = non_ascii_str_enc_binary
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16BE') }
  end

  def test_ascii_str_utf8_to_ascii
    str = ascii_str_enc_utf8
    assert_equal(Encoding::ASCII, str.encode('ASCII').encoding)
  end

  def test_non_ascii_str_utf8_to_ascii
    str = non_ascii_str_enc_utf8
    assert_raise(Encoding::UndefinedConversionError) { str.encode('ASCII') }
  end

  def test_ascii_str_utf8_to_binary
    str = ascii_str_enc_utf8
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end

  def test_non_ascii_str_utf8_to_binary
    str = non_ascii_str_enc_utf8
    assert_raise(Encoding::UndefinedConversionError) { str.encode('BINARY') }
  end

  def test_ascii_str_utf8_to_utf8
    str = ascii_str_enc_utf8
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_8') }
  end

  def test_non_ascii_str_utf8_to_utf8
    str = non_ascii_str_enc_utf8
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_8') }
  end

  def test_ascii_str_utf8_to_eucjp
    str = ascii_str_enc_utf8
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_non_ascii_str_utf8_to_eucjp
    str = non_ascii_str_enc_utf8
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_ascii_str_utf8_to_sjis
    str = ascii_str_enc_utf8
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_non_ascii_str_utf8_to_sjis
    str = non_ascii_str_enc_utf8
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_ascii_str_utf8_to_utf16le
    str = ascii_str_enc_utf8
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16LE') }
  end

  def test_non_ascii_str_utf8_to_utf16le
    str = non_ascii_str_enc_utf8
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16LE') }
  end

  def test_ascii_str_utf8_to_utf16be
    str = ascii_str_enc_utf8
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16BE') }
  end

  def test_non_ascii_str_utf8_to_utf16be
    str = non_ascii_str_enc_utf8
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16BE') }
  end

  def test_ascii_str_eucjp_to_ascii
    str = ascii_str_enc_eucjp
    assert_equal(Encoding::ASCII, str.encode('ASCII').encoding)
  end

  def test_non_ascii_str_eucjp_to_ascii
    str = non_ascii_str_enc_eucjp
    assert_raise(Encoding::UndefinedConversionError) { str.encode('ASCII') }
  end

  def test_ascii_str_eucjp_to_binary
    str = ascii_str_enc_eucjp
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end

  def test_non_ascii_str_eucjp_to_binary
    str = non_ascii_str_enc_eucjp
    assert_raise(Encoding::UndefinedConversionError) { str.encode('BINARY') }
  end

  def test_ascii_str_eucjp_to_utf8
    str = ascii_str_enc_eucjp
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_8') }
  end

  def test_non_ascii_str_eucjp_to_utf8
    str = non_ascii_str_enc_eucjp
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_8') }
  end

  def test_ascii_str_eucjp_to_eucjp
    str = ascii_str_enc_eucjp
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_non_ascii_str_eucjp_to_eucjp
    str = non_ascii_str_enc_eucjp
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_ascii_str_eucjp_to_sjis
    str = ascii_str_enc_eucjp
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_non_ascii_str_eucjp_to_sjis
    str = non_ascii_str_enc_eucjp
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_ascii_str_eucjp_to_utf16le
    str = ascii_str_enc_eucjp
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16LE') }
  end

  def test_non_ascii_str_eucjp_to_utf16le
    str = non_ascii_str_enc_eucjp
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16LE') }
  end

  def test_ascii_str_eucjp_to_utf16be
    str = ascii_str_enc_eucjp
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16BE') }
  end

  def test_non_ascii_str_eucjp_to_utf16be
    str = non_ascii_str_enc_eucjp
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16BE') }
  end

  def test_ascii_str_sjis_to_ascii
    str = ascii_str_enc_sjis
    assert_equal(Encoding::ASCII, str.encode('ASCII').encoding)
  end

  def test_non_ascii_str_sjis_to_ascii
    str = non_ascii_str_enc_sjis
    assert_raise(Encoding::UndefinedConversionError) { str.encode('ASCII') }
  end

  def test_ascii_str_sjis_to_binary
    str = ascii_str_enc_sjis
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end

  def test_non_ascii_str_sjis_to_binary
    str = non_ascii_str_enc_sjis
    assert_raise(Encoding::UndefinedConversionError) { str.encode('BINARY') }
  end

  def test_ascii_str_sjis_to_utf8
    str = ascii_str_enc_sjis
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_8') }
  end

  def test_non_ascii_str_sjis_to_utf8
    str = non_ascii_str_enc_sjis
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_8') }
  end

  def test_ascii_str_sjis_to_eucjp
    str = ascii_str_enc_sjis
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_non_ascii_str_sjis_to_eucjp
    str = non_ascii_str_enc_sjis
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_ascii_str_sjis_to_sjis
    str = ascii_str_enc_sjis
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_non_ascii_str_sjis_to_sjis
    str = non_ascii_str_enc_sjis
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_ascii_str_sjis_to_utf16le
    str = ascii_str_enc_sjis
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16LE') }
  end

  def test_non_ascii_str_sjis_to_utf16le
    str = non_ascii_str_enc_sjis
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16LE') }
  end

  def test_ascii_str_sjis_to_utf16be
    str = ascii_str_enc_sjis
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16BE') }
  end

  def test_non_ascii_str_sjis_to_utf16be
    str = non_ascii_str_enc_sjis
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16BE') }
  end

  def test_ascii_str_utf16le_to_ascii
    str = ascii_str_enc_utf16le
    assert_equal(Encoding::ASCII, str.encode('ASCII').encoding)
  end

  def test_non_ascii_str_utf16le_to_ascii
    str = non_ascii_str_enc_utf16le
    assert_raise(Encoding::UndefinedConversionError) { str.encode('ASCII') }
  end

  def test_ascii_str_utf16le_to_binary
    str = ascii_str_enc_utf16le
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end

  def test_non_ascii_str_utf16le_to_binary
    str = non_ascii_str_enc_utf16le
    assert_raise(Encoding::UndefinedConversionError) { str.encode('BINARY') }
  end

  def test_ascii_str_utf16le_to_utf8
    str = ascii_str_enc_utf16le
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_8') }
  end

  def test_non_ascii_str_utf16le_to_utf8
    str = non_ascii_str_enc_utf16le
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_8') }
  end

  def test_ascii_str_utf16le_to_eucjp
    str = ascii_str_enc_utf16le
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_non_ascii_str_utf16le_to_eucjp
    str = non_ascii_str_enc_utf16le
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_ascii_str_utf16le_to_sjis
    str = ascii_str_enc_utf16le
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_non_ascii_str_utf16le_to_sjis
    str = non_ascii_str_enc_utf16le
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_ascii_str_utf16le_to_utf16le
    str = ascii_str_enc_utf16le
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16LE') }
  end

  def test_non_ascii_str_utf16le_to_utf16le
    str = non_ascii_str_enc_utf16le
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16LE') }
  end

  def test_ascii_str_utf16le_to_utf16be
    str = ascii_str_enc_utf16le
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16BE') }
  end

  def test_non_ascii_str_utf16le_to_utf16be
    str = non_ascii_str_enc_utf16le
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16BE') }
  end

  def test_ascii_str_utf16be_to_ascii
    str = ascii_str_enc_utf16be
    assert_equal(Encoding::ASCII, str.encode('ASCII').encoding)
  end

  def test_non_ascii_str_utf16be_to_ascii
    str = non_ascii_str_enc_utf16be
    assert_raise(Encoding::UndefinedConversionError) { str.encode('ASCII') }
  end

  def test_ascii_str_utf16be_to_binary
    str = ascii_str_enc_utf16be
    assert_equal(Encoding::BINARY, str.encode('BINARY').encoding)
  end

  def test_non_ascii_str_utf16be_to_binary
    str = non_ascii_str_enc_utf16be
    assert_raise(Encoding::UndefinedConversionError) { str.encode('BINARY') }
  end

  def test_ascii_str_utf16be_to_utf8
    str = ascii_str_enc_utf16be
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_8') }
  end

  def test_non_ascii_str_utf16be_to_utf8
    str = non_ascii_str_enc_utf16be
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_8') }
  end

  def test_ascii_str_utf16be_to_eucjp
    str = ascii_str_enc_utf16be
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_non_ascii_str_utf16be_to_eucjp
    str = non_ascii_str_enc_utf16be
    assert_equal(Encoding::EUCJP, str.encode('EUCJP').encoding)
  end

  def test_ascii_str_utf16be_to_sjis
    str = ascii_str_enc_utf16be
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_non_ascii_str_utf16be_to_sjis
    str = non_ascii_str_enc_utf16be
    assert_equal(Encoding::SJIS, str.encode('SJIS').encoding)
  end

  def test_ascii_str_utf16be_to_utf16le
    str = ascii_str_enc_utf16be
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16LE') }
  end

  def test_non_ascii_str_utf16be_to_utf16le
    str = non_ascii_str_enc_utf16be
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16LE') }
  end

  def test_ascii_str_utf16be_to_utf16be
    str = ascii_str_enc_utf16be
    assert_raise(Encoding::ConverterNotFoundError){ str.encode('UTF_16BE') }
  end

  def test_non_ascii_str_utf16be_to_utf16be
    str = non_ascii_str_enc_utf16be
    assert_raise(Encoding::ConverterNotFoundError) { str.encode('UTF_16BE') }
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
