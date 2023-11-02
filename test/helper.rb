# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'bundler'

begin
  Bundler.setup(:default, :development)
rescue Bundler::BundlerError => e
  warn e.message
  warn "Run `bundle install` to install missing gems"
  exit e.status_code
end
require 'minitest/autorun'

$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'writeexcel'

class Minitest::Test
  ###############################################################################
  #
  # Unpack the binary data into a format suitable for printing in tests.
  #
  def unpack_record(data)
    data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
  end

  # expected : existing file path
  # target   : io (ex) string io object where stored data.
  def compare_file(expected, target)
    # target is StringIO object.
    result =
      ruby_18 { target.string } ||
      ruby_19 { target.string.force_encoding('BINARY') }
    assert_equal(
      File.binread(expected),
      result,
      "#{File.basename(expected)} doesn't match."
    )
  end
end
