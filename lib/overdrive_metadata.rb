require 'htmlentities'
require 'marc'
require 'sanitize'
require 'spreadsheet'

##
# Class to generate marc records from Overdrive provided metadata spreadsheet

class OverdriveMetadata
  	VERSION = '1.0.0'

  	attr_reader :records

    GMD      = '[electronic resource]'
    ISBN_AUD = '(sound recording : OverDrive Audio Book)'
    ACCESS   = 'Mode of access: World Wide Web.'
    REQUIRE  = 'Requires OverDrive Media Console'
    # INV_SUMM = '&lt;p&gt;'
    DISCLAIM = 'Record generated from Overdrive metadata spreadsheet.'
    URL_MSG  = 'Click to download this audiobook.'
    EXC_MSG  = 'Excerpt.'
    DOWN_SH  = 'Downloadable audiobooks.'

    READ_ERR = 'Read error, check file path or try resaving file as .xls (not xml)'
    TITL_ERR = 'Title field is missing from data row'

    # add option for config. file in future 
    HEADERS = {
      :oclc      => 19,
      :date      => 12,
      :time      => 21,
      :isbn      => 1,
      :author    => 4,
      :title     => 2,
      :place     => 11,
      :publisher => 3,
      :filesize  => 8,
      :reader    => 14,
      :title_src => 13,
      :summary   => 15,
      :subjects  => 5,
      :download  => 7,
      :excerpt   => 16,
      :cover     => 17,
      :thumb     => 18,
    }

  	def initialize(metadata_file)
  		begin	
  			@metadata = Spreadsheet.open(metadata_file).worksheet 0
  			@coder = HTMLEntities.new
        @records = []
  		rescue Exception => ex
  			raise READ_ERR
  		end
  	end

  	def map
  		@metadata.each do |row|
    		@records << create_record(row)
  		end
  	end

  	def create_record(data)
  		record = MARC::Record.new

      data[HEADERS[:author]]  = @coder.decode(data[HEADERS[:author]])
      data[HEADERS[:title]]   = @coder.decode(data[HEADERS[:title]])
      data[HEADERS[:reader]]  = @coder.decode(data[HEADERS[:reader]])
      data[HEADERS[:summary]] = Sanitize.clean(@coder.decode(data[HEADERS[:summary]]))
  		
      date             = data[HEADERS[:date]]
      year             = month = day = ''
      if date.match(/\d{2}\/\d{2}\/\d{4}/)
        month, day, year = date.split '/'
      end
      year             = date.match(/\d{4}/).to_s # Fall-back
      isbn             = data[HEADERS[:isbn]]
      time             = data[HEADERS[:time]]
      hr, mn, sc       = time.split ':'
      author           = normalize_author data[HEADERS[:author]]
      title            = data[HEADERS[:title]]
      reader           = data[HEADERS[:reader]]
      summary          = data[HEADERS[:summary]]
      subjects         = data[HEADERS[:subjects]].split ','

      # leader - accept hash in future
      ldr = record.leader
      ldr[5]  = 'n'
      ldr[6]  = 'i'
      ldr[7]  = 'm'
      ldr[17] = '4'
      ldr[18] = 'a'
      
      fields = []
  		fields << make_control_field('001', data[HEADERS[:oclc]]) # oclc no.
      fields << make_control_field('006', 'm        h        ')
      fields << make_control_field('007', 'sz usnnnn   ed')
      fields << make_control_field('007', 'cr nna        ')
      fields << make_control_field('008', make_fixed_field(year, month, day))

  		fields << make_data_field('020', ' ', ' ', isbn + ' ' + ISBN_AUD) unless isbn.empty?
      fields << make_data_field('037', ' ', ' ', 'OverDrive, Inc.', 'b')
      fields << make_source('JTH')
  		
  		fields << make_data_field('100', '1', ' ', author)
      fields << make_title(title, data[HEADERS[:author]])
      fields << make_publication(data[HEADERS[:place]], data[HEADERS[:publisher]], year)
      fields << make_physical(hr, mn)

      fields << make_data_field('306', ' ', ' ', hr + mn + sc)
      fields << make_data_field('538', ' ', ' ', ACCESS)
      fields << make_data_field('538', ' ', ' ', REQUIRE + ' (file size: ' + data[HEADERS[:filesize]] + ' KB).')
      fields << make_data_field('511', '0', ' ', 'Read by ' + reader + '.') unless reader.empty? 
      fields << make_data_field('520', ' ', ' ', summary) unless summary.match(/^#+$/)

      fields << make_data_field('500', ' ', ' ', 'Title from: ' + data[HEADERS[:title_src]] + '.')
      fields << make_data_field('500', ' ', ' ', 'Unabridged.')
      fields << make_data_field('500', ' ', ' ', 'Duration: ' + hr + ' hr., ' + mn + ' min.')

      subjects.each { |s| fields << make_subject(@coder.decode(s)) }
      fields << make_dlc_sh

      added = normalize_author reader
      fields << make_data_field('700', '1', ' ', added)

      fields << make_link(data[HEADERS[:download]], URL_MSG)
      fields << make_link(data[HEADERS[:excerpt]], EXC_MSG)
      fields << make_img_link(title, data[HEADERS[:cover]], data[HEADERS[:thumb]])

      fields << make_data_field('907', ' ', ' ', 'ER')
      fields << make_data_field('991', ' ', ' ', DISCLAIM)

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

    def make_fixed_field(year, month, day)
      return nil if year.empty?
      fixed_field = '      s        xxunnnn s           eng d'
      if month and day
        fixed_field[0..5] = year[2..3] + month + day
        fixed_field[7..10] = year
      else
        fixed_field[7..10] = year
      end
      return fixed_field
    end

    def make_source(agency)
      src_f = make_data_field('040', ' ', ' ', agency)
      append_subfield src_f, 'c', agency
      return src_f
    end

  	def make_title(title, sor)
      raise TITL_ERR if title.empty?
  		t_ind1  = sor.empty? ? '0' : '1'
  		t_ind2  = non_filing_characters title
  		title_f = make_data_field('245', t_ind1, t_ind2, title)
      append_subfield title_f, 'h', GMD + ' /'
  		unless sor.empty?
  			append_subfield title_f, 'c', 'by ' + sor + '.'
      else
        title_f['h'].gsub!(/\s+\/$/, '.')
  		end
  		return title_f
  	end

    def make_publication(place, publisher, year)
      return nil if place.empty? or publisher.empty? or year.empty?
      pub_f = make_data_field('260', ' ', ' ', place + ' :')
      append_subfield pub_f, 'b', publisher + ','
      append_subfield pub_f, 'c', year + '.'
      return pub_f
    end

    def make_physical(hours, minutes)
      phys_f = make_data_field('300', ' ', ' ', '1 sound file (ca. ' + hours + ' hr., ' + minutes + ' min.) :')
      append_subfield phys_f, 'b', 'digital'
      return phys_f
    end

    def make_subject(subject)
      subj_f  = make_data_field('655', ' ', '7', subject.strip + '.')
      append_subfield subj_f, '2', 'local'
      return subj_f
    end

    # cludgy, fix later ...
    def make_dlc_sh
      dlc_f = make_data_field('655', ' ', '7', DOWN_SH)
      append_subfield dlc_f, '2', 'local'
      return dlc_f
    end

    def make_link(url, message)
      link_f = make_data_field('856', '4', '0', url)
      append_subfield link_f, 'y', message
      return link_f
    end

    def make_img_link(title, cover, thumb)
      img_f = make_data_field('856', '4', '2', cover)
      append_subfield img_f, 'y', "<img class='scl_mwthumb' src='" + thumb + "' alt='Artwork for this title - " + title + "' />"
      return img_f
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
  			'4'
  		when title.match(/^An /)
  			'3'
  		when title.match(/^A /)
  			'2'		
  		else
  			'0'
  		end
  	end

end
