= overdrive_metadata

http://www.libcode.net

== DESCRIPTION:

Generate marc records from Overdrive provided metadata spreadsheets.

== FEATURES/PROBLEMS:

Have yet to see a Kindle eBook sample - may require tinkering.

== SYNOPSIS:

require 'overdrive_metadata'
records = OverdriveMetadata.new('spreadsheets/111111.xls')
puts "R: " + records.size.to_s # print number of records generated to console
w = MARC::Writer.new('generated.mrc')
records.each do |r|
  begin
    w.write r
  rescue
    puts "FAILED: " + r['245']['a']
  end
end
w.close

== REQUIREMENTS:

htmlentities
marc
sanitize
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
