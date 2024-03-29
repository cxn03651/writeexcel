# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'writeexcel/version'

Gem::Specification.new do |gem|
  gem.name          = "writeexcel"
  gem.version       = Writeexcel::VERSION
  gem.authors       = ["Hideo NAKAMURA"]
  gem.email         = ["nakamura.hideo@gmail.com"]
  gem.description   = "Multiple worksheets can be added to a workbook and formatting can be applied to cells. Text, numbers, formulas, hyperlinks and images can be written to the cells."
  gem.summary       = "Write to a cross-platform Excel binary file."
  gem.homepage      = "http://github.com/cxn03651/writeexcel#readme"
  gem.license       = 'MIT'

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ["lib"]
  gem.required_ruby_version = '>= 2.4.0'
  gem.add_development_dependency 'minitest'
  gem.add_runtime_dependency 'racc' if RUBY_VERSION >= '3.3'
  gem.add_runtime_dependency 'nkf'
  gem.add_development_dependency 'rake'
  gem.extra_rdoc_files = [
    "LICENSE.txt",
    "README.rdoc"
  ]
  gem.add_development_dependency 'simplecov'
end
