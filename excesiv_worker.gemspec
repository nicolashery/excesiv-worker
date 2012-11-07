$:.push File.expand_path("../lib", __FILE__)
require 'excesiv/version'

Gem::Specification.new do |s|
  s.name          = 'excesiv_worker'
  s.version       = Excesiv::VERSION
  s.authors       = ['Nicolas Hery']
  s.email         = 'nicolahery@gmail.com'
  s.summary       = 'Worker module of Excesiv, Excel file generator and reader'
  s.homepage      = 'http://github.com/nicolahery/excesiv-worker'

  s.files         = `git ls-files`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ['lib']

  s.add_dependency 'mongo', '~> 1.7'
end