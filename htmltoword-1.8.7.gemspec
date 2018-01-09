# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'htmltoword/version'

Gem::Specification.new do |spec|
  spec.name = 'htmltoword-1.8.7'
  spec.version = Htmltoword::VERSION
  spec.authors = ['Nicholas Frandsen, Cristina Matonte', 'Ayoub Khobaklatte']
  spec.email = ['nick.rowe.frandsen@gmail.com, anitsirc1@gmail.com', 'ayoub.khobalatte@gmail.com']
  spec.description = %q{Convert html to word docx document.}
  spec.summary = %q{This simple gem allows you to create MS Word docx documents from simple html documents. This makes it easy to create dynamic reports and forms that can be downloaded by your users as simple MS Word docx files.}
  spec.homepage = 'http://github.com/unwire/htmltoword'
  spec.license = 'MIT'

  spec.files = Dir.glob('{lib}/**/*.rb') + Dir.glob('{templates,xslt}/*') + %w{ README.md Rakefile }
  spec.executables = spec.files.grep(%r{^bin/}) {|f| File.basename(f)}
  spec.test_files = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ['lib']

  spec.add_dependency 'actionpack', '~> 2.3'
  spec.add_dependency 'nokogiri', '1.5.5'
  spec.add_dependency 'rubyzip', '0.9.9'

  spec.add_development_dependency 'rspec'
  spec.add_development_dependency 'bundler'
  spec.add_development_dependency 'rake'

end
