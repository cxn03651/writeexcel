# -*- coding: utf-8 -*-
##########################################################################
# test_23_note.rb
#
# Tests for some of the internal method used to write the NOTE record that
# is used in cell comments.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#
#########################################################################
require 'helper'
require 'stringio'

class TC_note < Test::Unit::TestCase

  def setup
    @workbook   = WriteExcel.new(StringIO.new)
    @worksheet  = @workbook.add_worksheet
  end

  def test_blank_author_name
    comment = Writeexcel::Worksheet::Comment.new(@worksheet, 2, 0, 'Test')
    obj_id   = 1

    caption = sprintf(" \tnote_record")
    target  = %w(
        1C 00 0C 00 02 00 00 00 00 00 01 00 00 00 00 00
    ).join(' ')
    result = unpack_record(comment.note_record(obj_id))
    assert_equal(target, result, caption)
  end

  def test_defined_author_name
    comment = Writeexcel::Worksheet::Comment.new(@worksheet, 2, 0,'Test', :author => 'Username')
    obj_id   = 1

    caption = sprintf(" \tstore_note")
    target  = %w(
        1C 00 14 00 02 00 00 00 00 00 01 00 08 00 00 55
        73 65 72 6E 61 6D 65 00
    ).join(' ')
    result = unpack_record(comment.note_record(obj_id))
    assert_equal(target, result, caption)
  end

  ###############################################################################
  #
  # Unpack the binary data into a format suitable for printing in tests.
  #
  def unpack_record(data)
    data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
  end

end
