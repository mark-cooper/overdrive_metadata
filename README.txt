= overdrive_metadata

http://www.libcode.net
https://github.com/mark-cooper/overdrive_metadata

== DESCRIPTION:

Generate marc records from Overdrive provided metadata spreadsheets.

== FEATURES/PROBLEMS:

Important! Open and save the spreadsheet provided by Overdrive as a plain .xls (2003).
By default they come as an XML XLS which is unreadable by the Ruby spreadsheet library.

Much faster than previous versions -- no batch merging.
Fields are appended to a single record for rows with matching content urls.
Updated to account for ebook formats.
Now agency code must be passed in as second argument and headers are assumed to be present.

== SYNOPSIS:

require 'overdrive_metadata'

o = OverdriveMetadata.new('spreadsheets/111111.xls', 'JTH')

OR

o = OverdriveMetadata.new('spreadsheets/111111.xls', 'JTH', false) # if no header

records = o.map # this must be called to process the rows

# count of spreadsheet rows processed
puts "Number of fields read: #{o.count.to_s}"

# print number of records generated to console
puts "Number of records: #{records.size.to_s}"

w = MARC::Writer.new('generated.mrc')

records.each { |r| w.write r }

w.close

== REQUIREMENTS:

marc
spreadsheet

== INSTALL:

sudo gem install overdrive_metadata

== LICENSE:

(The MIT License)

Copyright (c) 2011 Mark Cooper

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
'Software'), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
