require 'spreadsheet'

class OverdriveMetadata
  	VERSION = '1.0.0'

  	attr_reader :records

  	def initialize(metadata_file)
  		begin	
  			md = Spreadsheet.open(metadata_file).worksheet 0
  			@records = []
  		rescue Exception => ex
  			raise 'Spreadsheet read error, try resaving file as .xls (not xml)'
  		end
  	end

  	def map(metadata)
  		metadata.each do |row|
    		@records << create_record(row)
  		end
  	end

  	def create_record(data)
  		record = MARC::Record.new
  		fields = []
  		# leader
  		fields << make_control_field('001', data[19]) # oclc no.
  		# 007 ...
  		# 008 ...
  		fields << make_data_field('020', ' ', ' ', data[1]) # isbn no.

  		fields << make_data_field('037', ' ', ' ', 'OverDrive, Inc.', 'b')
  		
  		author = normalize_author data[4]
  		fields << make_data_field('100', '1', '', author)

		fields << make_title(data[2], data[4])

  		fields.compact! # remove nil fields
  		fields.each { |f| record.append f }

  		return record
  	end

  	def make_control_field(tag, value)
  		return nil if value.empty?
  		return MARC::ControlField.new(tag, value)
  	end

  	def make_data_field(tag, ind1, ind2, value, code = 'a')
  		return nil if value.empty?
  		return MARC::DataField.new(tag, ind1, ind2, [code, value])
  	end

  	def make_title(title, sor)
  		title   = sor.empty? ? title + '.' : title + ' /'
		t_ind1  = sor.nil? ? '0' : '1'
		t_ind2  = non_filing_characters title
		title_f = make_data_field('245', t_ind1, t_ind2, title)
		unless sor.empty?
			append_subfield title_f, 'c', 'by ' + sor + '.'
		end
		return title_f
  	end

  	def append_subfield(field, code, value)
  		return field.append(MARC::Subfield.new(code, value))
  	end

  	def normalize_author(author)
		return author if author.empty?
		names    = author.split ' '
		surname  = names.last + ', '
		fullname = surname + names[0 .. names.length - 2].join(' ')
		fullname += '.' unless fullname[-1] == '.'
		return fullname
	end

	def non_filing_characters(title)
		return case
		when title.match(/^The /)
			4
		when title.match(/^An /)
			3
		when title.match(/^A /)
			2		
		else
			0
		end
	end

end
