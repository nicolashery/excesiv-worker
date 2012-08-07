# Wrapper around selected classes of the Apache POI Excel library
require 'bundler/setup'
require 'java'
require 'require_all'
# Require Apache POI jars
require_rel 'lib/poi/*.jar'

module Poi
  XSSFWorkbook = Java::OrgApachePoiXssfUsermodel::XSSFWorkbook
  AreaReference = Java::OrgApachePoiSsUtil::AreaReference
  CELL_TYPE_NUMERIC = 0
  CELL_TYPE_STRING = 1
  CELL_TYPE_FORMULA = 2
  CELL_TYPE_BLANK = 3
  CELL_TYPE_BOOLEAN = 4
  CELL_TYPE_ERROR = 5
end 
