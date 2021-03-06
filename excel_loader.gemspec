# -*- encoding: utf-8 -*-
$LOAD_PATH.unshift File.expand_path('../lib', __FILE__)

Gem::Specification.new do |s|
  s.name        = "excel_loader"
  s.version     = "0.0.6"
  s.platform    = Gem::Platform::RUBY
  s.authors     = ["Jon Sarley"]
  s.email       = ["jsarley@softmodal.com"]
  s.homepage    = "http://softmodal.com"
  s.summary     = %q{Dumps and retrieves tabular data to and from an Excel file.}
  s.description     = %q{Dumps and retrieves tabular data to and from an Excel file.}

  s.add_dependency "spreadsheet"
  s.add_development_dependency "rspec"

  s.files         = `git ls-files`.split("\n")
  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
end