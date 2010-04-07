require 'helper'

class TC_BIFFWriter < Test::Unit::TestCase

  TEST_DIR    = File.expand_path(File.dirname(__FILE__))
  PERL_OUTDIR = File.join(TEST_DIR, 'perl_output')

  def setup
    t = Time.now.strftime("%Y%m%d")
    path = "temp#{t}-#{$$}-#{rand(0x100000000).to_s(36)}"
    @test_file  = File.join(Dir.tmpdir, path)
    @biff = BIFFWriter.new
    @ruby_file = @test_file
  end

  def test_append_no_error
    assert_nothing_raised{ @biff.append("World") }
  end

  def test_prepend_no_error
    assert_nothing_raised{ @biff.prepend("Hello") }
  end

  def test_data_added
    assert_nothing_raised{ @biff.append("Hello", "World") }
    data = ''
    while d = @biff.get_data
      data += d
    end
    assert_equal("HelloWorld", data, "Bad data contents")
    assert_equal(10, @biff.datasize, "Bad data size")
  end

  def test_data_prepended

    assert_nothing_raised{ @biff.append("Hello") }
    assert_nothing_raised{ @biff.prepend("World") }
    data = ''
    while d = @biff.get_data
      data += d
    end
    assert_equal("WorldHello", data, "Bad data contents")
    assert_equal(10, @biff.datasize, "Bad data size")
  end

  def test_store_bof_length
    assert_nothing_raised{ @biff.store_bof }
    assert_equal(20, @biff.datasize, "Bad data size after store_bof call")
  end

  def test_store_eof_length
    assert_nothing_raised{ @biff.store_eof }
    assert_equal(4, @biff.datasize, "Bad data size after store_eof call")
  end

  def test_datasize_mixed
    assert_nothing_raised{ @biff.append("Hello") }
    assert_nothing_raised{ @biff.prepend("World") }
    assert_nothing_raised{ @biff.store_bof }
    assert_nothing_raised{ @biff.store_eof }
    assert_equal(34, @biff.datasize, "Bad data size for mixed data")
  end

  def test_add_continue
    perl_file = "#{PERL_OUTDIR}/biff_add_continue_testdata"
    size = File.size(perl_file)
    @fh = File.new(@ruby_file,"w+")
    @fh.print(@biff.add_continue('testdata'))
    @fh.close
    rsize = File.size(@ruby_file)
    assert_equal(size,rsize,"File sizes not the same")
    compare_file(perl_file, @ruby_file)
  end

  def teardown
    @biff = nil
    File.delete(@ruby_file) if File.exist?(@ruby_file)
  end

  def compare_file(expected, target)
    fh_e = File.open(expected, "r")
    fh_t = File.open(target, "r")
    while true do
      e1 = fh_e.read(1)
      t1 = fh_t.read(1)
      if e1.nil?
        assert( t1.nil?, "#{expexted} is EOF but #{target} is NOT EOF.")
        break
      elsif t1.nil?
        assert( e1.nil?, '#{target} is EOF but #{expected} is NOT EOF.')
        break
      end
      assert_equal(e1, t1, sprintf(" #{expected} = '%s' but #{target} = '%s'", e1, t1))
      break
    end
    fh_e.close
    fh_t.close
  end


end
