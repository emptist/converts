# read from a folder of xlsx files with filter then combine into one
cej = require 'convert-excel-to-json'
fs = require 'fs'
xlsx = require 'json-as-xlsx'

# first rename all files as this neededRowName_otherwords.xlsx
read = (funcOpts) ->
  {sourceFile,jsonfilename,header,sheets,range,columnToKey} = funcOpts
  jsonContent = cej {sourceFile,header,sheets,range,columnToKey}
  rowName = sourceFile.split('_')[0]
  reg = new RegExp(rowName)
  rowContents = (each for each in jsonContent when reg.test(each))
  {rowName,rowContents}

write = (funcOpts) ->
