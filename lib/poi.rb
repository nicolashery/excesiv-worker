# Wrapper around selected classes of the Apache POI Excel library
require 'bundler/setup'
require 'java'

# Require Apache POI jars
require 'poi/commons-logging-1.1.jar'
require 'poi/dom4j-1.6.1.jar'
require 'poi/log4j-1.2.13.jar'
require 'poi/poi-3.8-20120326.jar'
require 'poi/poi-ooxml-3.8-20120326.jar'
require 'poi/poi-ooxml-schemas-3.8-20120326.jar'
require 'poi/stax-api-1.0.1.jar'
require 'poi/xmlbeans-2.3.0.jar'

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
