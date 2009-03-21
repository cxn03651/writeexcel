require 'rubygems'

spec = Gem::Specification.new do |s|
  s.name     = 'writeexcel'
  s.version  = '0.1.0'
  s.author   = 'Hideo Nakamura'
  s.email    = 'cxn03651@msj.biglobe.ne.jp'
  s.platform = 'Gem::Platform::RUBY'
  s.summary  = 'Write to a cross-platform Excel binary file.'
  s.files    = Dir['examples/**/*'] + Dir['lib/**/*.rb'] +
               Dir['[A-Z]*'] + Dir['test/**/*']
  s.require_path = 'lib'
  s.autorequire  = 'writeexcel'
  s.test_file    = 'test/ts_all.rb'
  s.has_rdoc     = false
  s.add_dependency('ruby-ole', '>=1.2.8.2')
end

if $0 == __FILE__
  Gem::manage_gems
  Gem::Builder.new(spec).build
end
