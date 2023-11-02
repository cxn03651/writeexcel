# -*- coding: utf-8 -*-
require 'helper'
require 'stringio'

class TC_BIFFWriter < Minitest::Test

  TEST_DIR    = File.expand_path(File.dirname(__FILE__))
  PERL_OUTDIR = File.join(TEST_DIR, 'perl_output')

  def setup
    @biff = BIFFWriter.new
    @ruby_file = StringIO.new
  end

  def test_data_added
    data = ''
    while d = @biff.get_data
      data += d
    end
    assert_equal("HelloWorld", data, "Bad data contents")
    assert_equal(10, @biff.datasize, "Bad data size")
  end

  def test_data_prepended
    data = ''
    while d = @biff.get_data
      data += d
    end
    assert_equal("WorldHello", data, "Bad data contents")
    assert_equal(10, @biff.datasize, "Bad data size")
  end

  def test_store_bof_length
    assert_equal(20, @biff.datasize, "Bad data size after store_bof call")
  end

  def test_store_eof_length
    assert_equal(4, @biff.datasize, "Bad data size after store_eof call")
  end

  def test_datasize_mixed
    assert_equal(34, @biff.datasize, "Bad data size for mixed data")
  end

  def test_add_continue
    perl_file = "#{PERL_OUTDIR}/biff_add_continue_testdata"
    size = File.size(perl_file)
    @ruby_file.print(@biff.add_continue('testdata'))
    rsize = @ruby_file.size
    assert_equal(size, rsize, "File sizes not the same")
    compare_file(perl_file, @ruby_file)
  end
end
