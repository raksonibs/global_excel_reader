# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'global_excel_reader/version'

Gem::Specification.new do |gem|
  gem.name          = "global_excel_reader"
  gem.version       = GlobalExcelReader::VERSION
  gem.authors       = ["Oskar Niburski"]
  gem.email         = ["oskarniburski@gmail.com"]
  gem.description   = %q{Read all excel data with Ruby and grab styles, data, etc.}
  gem.summary       = %q{Read all excel data with Ruby and grab styles, data, etc.}
  gem.homepage      = ""

  gem.add_dependency 'nokogiri'
  gem.add_dependency 'rubyzip'

  gem.add_development_dependency 'minitest', '>= 5.0'
  gem.add_development_dependency 'rake'
  gem.add_development_dependency 'pry'

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^test/})
  gem.require_paths = ["lib"]
end
