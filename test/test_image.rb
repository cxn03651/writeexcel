# -*- coding: utf-8 -*-
require 'helper'

class TestImage < Test::Unit::TestCase
  def setup
    @test_file = StringIO.new
    @workbook = WriteExcel.new(@test_file)
    @worksheet = @workbook.add_worksheet('test', 0)
  end

  def test_import_image_is_created_by_adobe_photoshop
    image = Writeexcel::Image.new(@worksheet, 0, 0, "test/republic_ps.jpg")
    image.import
    assert_equal(image.height, 120)
    assert_equal(image.width, 120)
    assert_equal(image.size, 24972)
  end
end
