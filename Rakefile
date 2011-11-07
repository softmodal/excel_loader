require "rake"
require "rake/clean"
require "rake/gempackagetask"
require "rake/rdoctask"
require "rake/testtask"
require "spec/rake/spectask"

spec = Gem::Specification.new do |s| 
  s.name = "excel_loader"
  s.version = "0.0.4"
  s.author = "Jon Sarley"
  s.email = "jsarley@softmodal.com"
  s.homepage = "http://www.softmodal.com"
  s.platform = Gem::Platform::RUBY
  s.summary = "Dumps and retrieves tabular data to and from an Excel file."
  s.files = Dir["{bin,lib}/**/*"]
  s.require_path = "lib"
  s.has_rdoc = true
  s.add_dependency("spreadsheet")
  s.add_dependency("rspec")
end
 
Rake::GemPackageTask.new(spec) do |package|
  package.gem_spec = spec
end