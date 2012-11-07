# Use this file to launch from this directory with:
# jruby worker.rb

# Add the lib directory to the Ruby load path
$:.push File.expand_path("../lib", __FILE__)

require 'excesiv'

Excesiv::Worker.new.run