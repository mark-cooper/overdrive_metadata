require 'spreadsheet'

class OverdriveMetadata
  	VERSION = '1.0.0'

  	attr_reader :records

    HEADERS = {
      :oclc => 19,
      :date => 12,
      :isbn => 1,
      :author => 4,
      :title => 2,
    }

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
  		# leader
      ldr = record.leader
      
      fields = []
  		fields << make_control_field('001', data[HEADERS[:oclc]]) # oclc no.
      fields << make_control_field('006', 'm        h        ')
      fields << make_control_field('007', 'sz usnnnn   ed')
      fields << make_control_field('007', 'cr nna        ')
      fields << make_control_field('008', '      s        xxunnnn s           eng d')
      date = HEADERS[:date]
  		# 007 ...
  		# 008 ...
  		fields << make_data_field('020', ' ', ' ', data[HEADERS[:isbn]]) # isbn no.

  		fields << make_data_field('037', ' ', ' ', 'OverDrive, Inc.', 'b')
  		
  		author = normalize_author data[HEADERS[:author]]
  		fields << make_data_field('100', '1', '', author)

		  fields << make_title(data[HEADERS[:title]], data[HEADERS[:author]])

      fields << make_data_field('907', ' ', ' ', 'ER')
      fields << make_data_field('991', ' ', ' ', 'Generated from overdrive metadata spreadsheet.')

  		fields.compact! # remove nil fields
  		fields.each { |f| record.append f } # add fields to record

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

    def make_fixed_field(date)
      fixed_field = '      s        xxunnnn s           eng d'
      if date.match(/\d{2}\/\d{2}\/\d{4}/)
        m, d, y = date.split '/'
        fixed_field[0..5] = y[2..3] + m + d
        fixed_field[7..11] = y
      elsif date = date.match(/\d{4}/)
        fixed_field[7..11] = date.to_s
      end
      return fixed_field
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
