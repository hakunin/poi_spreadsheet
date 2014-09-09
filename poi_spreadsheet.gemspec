lib = File.expand_path('../lib/', __FILE__)
$:.unshift lib unless $:.include?(lib)

Gem::Specification.new do |s|
  s.name        = "poi_spreadsheet"
  s.version     = "0.0.3"
  s.date        = '2014-09-09'
  s.summary     = "Modify excel sheets using the powerfull Apache POI"
  s.description = "There are quite a few pure ruby gems to deal with excel files with various degree of success. This gem's approach is to use powerfull software like Apache POI and simply provide ruby like interface."
  s.authors     = ["Michal Hantl"]
  s.email       = 'michal@hantl.cz'
  s.files       = ["lib/poi_spreadsheet.rb"]
  s.homepage    = 'http://github.com/hakunin/poi_spreadsheet'
  s.license       = 'MIT'

  s.add_runtime_dependency "rjb", ["~> 1.4"]

  s.require_path = 'lib'
  s.files        = Dir.glob("{apache,lib}/**/*") + [
      "poi_spreadsheet.gemspec",
      "README.md"
    ]
end
