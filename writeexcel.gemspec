require 'rubygems'

version = '0.2.2'

spec = Gem::Specification.new do |s|
  s.name      = 'writeexcel'
  s.version   = version
  s.author    = 'Hideo NAKAMURA'
  s.email     = 'cxn03651@msj.biglobe.ne.jp'
  s.summary   = 'Write to a cross-platform Excel binary file.'
  s.files     = Dir['examples/**/*'] + Dir['lib/**/*.rb'] +
                Dir['[A-Z]*'] + Dir['test/**/*']
  s.require_path = 'lib'
  s.test_file    = 'test/ts_all.rb'
  s.has_rdoc     = false
  s.required_ruby_version = '>=1.8'
end
