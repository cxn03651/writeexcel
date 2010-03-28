$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../charts"

require "test/unit"
require 'writeexcel'
require 'chartex'

class TestChartex < Test::Unit::TestCase
  TEST_DIR    = File.expand_path(File.dirname(__FILE__))
  PERL_OUTDIR = File.join(TEST_DIR, 'perl_output')
  EXCEL_OUTDIR = File.join(TEST_DIR, 'excelfile')

  def setup
    @chartex = Chartex.new
  end

  def test_set_file
    filename = 'filename'
    @chartex.set_file(filename)
    assert_equal(filename, @chartex.file)
  end

  def test_get_workbook
    files = %w(1 2 3 4 5).collect { |i| "#{EXCEL_OUTDIR}/Chart#{i}.xls" }
    datas = %w(1 2 3 4 5).collect { |i| "#{PERL_OUTDIR}/Chart#{i}.xls.data" }

    (0...files.size).each do |i|
      @chartex.set_file(files[i])
      workbook = @chartex.get_workbook
      assert(workbook.kind_of?(OLEStorageLitePPS))
      expected = File.open(datas[i], 'rb') { |f| f.read }
      assert_equal(expected, workbook.data, "#{files[i]} failed.")
    end
  end
end
