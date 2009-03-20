require 'rubygems'
require 'rake'

SPEC = Gem::Specificaton.new do |s|
  s.name     = 'WriteExcel'
  s.version  = '0.1.0'
  s.author   = 'Hideo Nakamura'
  s.email    = 'cxn03651@msj.biglobe.ne.jp'
  s.platform = 'Gem::Platform::RUBY'
  s.summary  = 'Write to a cross-platform Excel binary file.'
  s.files    = FileList['examples/**/*', 'lib/**/*.rb', '{A-Z]*', 'test/**/*'].to_a
  s.require_path = 'lib'
  s.autorequire  = 'writeexcel'
  s.test_file    = 'test/ts_all.rb'
  s.has_rdoc     = false
  s.add_dependency('ruby-ole', '>=1.2.8.2')
end
