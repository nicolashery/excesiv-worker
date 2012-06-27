# Wrapper around selected classes of the Apache POI Excel library
require 'bundler/setup'
require 'java'
require 'require_all'
# Require Apache POI jars
require_rel 'lib/poi/*.jar'

module Poi
  XSSFWorkbook = Java::OrgApachePoiXssfUsermodel::XSSFWorkbook
end