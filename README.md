Overdrive Metadata Converter
============================

https://github.com/mark-cooper/overdrive_metadata

DESCRIPTION
-----------

Generate marc records from Overdrive provided metadata spreadsheets.

FEATURES/PROBLEMS
-----------------

Open and save the spreadsheet provided by Overdrive as a plain .xls (2003). By default they come as an XML XLS which is unreadable by the Ruby spreadsheet library.

- Much faster than previous versions -- no batch merging.
- Fields are appended to a single record for rows with matching content urls.
- Updated to account for ebook formats.
- Now agency code must be passed in as second argument and headers are assumed to be present.

SYNOPSIS
--------

Add the gem to a script:

    require 'overdrive_metadata'
    o = OverdriveMetadata.new('spreadsheets/111111.xls', 'JTH')
    # OR
    o = OverdriveMetadata.new('spreadsheets/111111.xls', 'JTH', false) # if no header

    records = o.map # this must be called to process the rows

    # count of spreadsheet rows processed
    puts "Number of fields read: #{o.count.to_s}"

    # print number of records generated to console
    puts "Number of records: #{records.size.to_s}"

    w = MARC::Writer.new('generated.mrc')
    records.each { |r| w.write r }
    w.close

REQUIREMENTS
------------

marc
spreadsheet

INSTALL
-------

sudo gem install overdrive_metadata

License & Authors
-----------------
- Author:: Mark Cooper

GPL v3