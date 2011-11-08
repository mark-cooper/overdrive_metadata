require 'htmlentities'
require 'marc'
require 'sanitize'
require 'spreadsheet'

##
# Class to generate marc records from Overdrive provided metadata spreadsheet
# Works for e-audiobooks only at present ...

class OverdriveMetadata
  	VERSION = '1.0.0'

  	attr_reader :records

    GMD      = '[electronic resource]'
    ISBN_AUD = '(sound recording : OverDrive Audio Book)'
    ACCESS   = 'Mode of access: World Wide Web.'
    DISCLAIM = 'Record generated from Overdrive metadata spreadsheet.'
    URL_MSG  = 'Click to download this audiobook.'
    DOWN_SH  = 'Downloadable audiobooks.'

    READ_ERR = 'Error, close file, check file path or try resaving file as .xls (not xml)'
    TITL_ERR = 'Title field is missing from data row'
    DATE_ERR = 'Date data not present for row'
    FIXF_ERR = 'Invalid fixed field created'

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
      :requires  => 10,
      :format    => 9,
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
        @coder    = HTMLEntities.new
        @records  = []
  		rescue Exception => ex
  			raise READ_ERR
  		end
  	end

  	def map
  		@metadata.each do |row|
    		@records << create_record(row)
  		end
      @records.compact!
      merge_by_isbn
  	end

    def merge_by_isbn
      isbns = Hash.new(0)
      @records.each do |record|
        isbns[record['020'].value] += 1 if record['020']
      end
      isbns.delete_if { |k,v| v < 2 }
      isbns.keys.each do |isbn|
        puts isbn
        rcds = @records.find_all { |r| r['020']['a'] == isbn if r['020'] }
        puts 'SIZE: ' + rcds.size.to_s
        raise 'Found invalid number of duplicate records: ' + isbn unless rcds.size == 2
        file_note = rcds[1].find { |f| f.tag == '500' and f['a'] =~ /OverDrive (WMA|MP3) Audiobook/ }
        excerpt   = rcds[1].find { |f| f.tag == '856' and f['y'] =~ /Excerpt/ }
        raise 'Unable to identify format note and excerpt: ' + isbn unless file_note and excerpt
        rcds[0].fields.insert(rcds[0].fields.index { |f| f.tag == '500' }, file_note)
        rcds[0].fields.insert(rcds[0].fields.index { |f| f.tag == '856' and f['y'] =~ /Excerpt/ }, excerpt)
        # rcds[0].append file_note
        # rcds[0].append excerpt
        # rcds[0].fields.sort_by! {|f| f.tag.to_i }
        @records.delete rcds[1]
      end
      @records
    end

  	def create_record(data)
  		record = MARC::Record.new
  		
      oclc             = data[HEADERS[:oclc]].to_s.empty? ? 'ovr' + make_id(data[HEADERS[:filesize]]) : 'ocn' + data[HEADERS[:oclc]]
      isbn             = data[HEADERS[:isbn]]
      date             = data[HEADERS[:date]]
      place            = data[HEADERS[:place]]
      publisher        = data[HEADERS[:publisher]]
      year             = month = day = ''
      if date.match(/\d{1,2}\/\d{1,2}\/\d{4}/)
        month, day, year = date.split '/'
      end
      year             = date.match(/\d{4}/).to_s # Fall-back
      time             = data[HEADERS[:time]]
      hr, mn, sc       = time.split ':'
      author           = @coder.decode(data[HEADERS[:author]])
      title            = @coder.decode(data[HEADERS[:title]])
      title_src        = data[HEADERS[:title_src]]
      reader           = @coder.decode(data[HEADERS[:reader]])
      requires         = data[HEADERS[:requires]]
      format           = data[HEADERS[:format]]
      filesize         = kb_to_mb(data[HEADERS[:filesize]])
      summary          = Sanitize.clean(@coder.decode(data[HEADERS[:summary]])).gsub(/\s{2}+/, '').strip
      subjects         = data[HEADERS[:subjects]].split ','
      litf             = subjects.include?('Fiction') ? 'f' : ' '
      download         = data[HEADERS[:download]]
      excerpt          = data[HEADERS[:excerpt]]
      thumb            = data[HEADERS[:thumb]]
      cover            = data[HEADERS[:cover]]

      # leader - accept hash in future
      ldr = record.leader
      ldr[5]  = 'n'
      ldr[6]  = 'i'
      ldr[7]  = 'm'
      ldr[17] = 'M'
      ldr[18] = 'a'
      
      begin
        fields = []
    		fields << make_control_field('001', oclc)
        fields << make_control_field('006', 'm        h        ')
        fields << make_control_field('007', 'sz usnnnnnnned')
        fields << make_control_field('007', 'cr nna   |||||')
        fields << make_control_field('008', make_fixed_field(year, month, day, litf))

    		fields << make_data_field('020', ' ', ' ', isbn + ' ' + ISBN_AUD) unless isbn.empty?
        fields << make_data_field('037', ' ', ' ', 'OverDrive, Inc.', 'b')
        fields << make_source('JTH')
    		
    		fields << make_data_field('100', '1', ' ', normalize_author(author))
        fields << make_title(title, author)
        fields << make_publication(place, publisher, year)
        fields << make_physical(hr, mn)

        fields << make_data_field('306', ' ', ' ', hr + mn + sc)
        fields << make_data_field('538', ' ', ' ', ACCESS)
        fields << make_data_field('538', ' ', ' ', 'Requires ' + requires + '.')
        fields << make_data_field('500', ' ', ' ', format + ' (file size: ' + filesize + ' MB).')
        fields << make_data_field('511', '0', ' ', 'Read by ' + reader + '.') unless reader.empty? 
        fields << make_data_field('520', ' ', ' ', summary) unless summary.match(/^#+$/)

        fields << make_data_field('500', ' ', ' ', 'Title from: ' + title_src + '.')
        fields << make_data_field('500', ' ', ' ', 'Unabridged.')
        fields << make_data_field('500', ' ', ' ', 'Duration: ' + hr + ' hr., ' + mn + ' min.')

        subjects.each { |s| fields << make_subject(@coder.decode(s)) }
        fields << make_dlc_sh

        fields << make_data_field('700', '1', ' ', normalize_author(reader))

        fields << make_link(download, URL_MSG)
        fields << make_link(excerpt, 'Excerpt (' + format + ').')
        fields << make_img_link(title, cover, thumb)

        fields << make_data_field('907', ' ', ' ', 'ER')
        fields << make_data_field('991', ' ', ' ', DISCLAIM)

    		fields.compact! # remove nil fields
    		fields.each { |f| record.append f } # add fields to record

    		return record
      rescue Exception => ex
        puts 'ERROR: ' + ex.message + ' for: ' + title
        return nil
      end
  	end

    ##
    # Generate a id no.

    def make_id(partial)
      return partial + Time.now.to_f.to_s.split('.')[1]
    end

  	def make_control_field(tag, value)
  		return nil if value.empty?
  		return MARC::ControlField.new(tag, value)
  	end

  	def make_data_field(tag, ind1, ind2, value, code = 'a')
  		return nil if value.empty?
  		return MARC::DataField.new(tag, ind1, ind2, [code, value])
  	end

    def make_fixed_field(year, month, day, litf = ' ')
      raise DATE_ERR if year.empty?
      fixed_field = '      s        xxunnnn s           eng d'
      unless month.empty? and day.empty?
        month = '0' + month if month.length == 1
        day   = '0' + day if day.length == 1
        fixed_field[0..5] = year[2..3] + month + day
        fixed_field[7..10] = year
      else
        fixed_field[7..10] = year
      end
      fixed_field[30] = litf
      raise FIXF_ERR unless fixed_field.length == 40
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
        value = sor[-1] == '.' ? "by #{sor}" : "by #{sor}."
  			append_subfield title_f, 'c', value
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
      return nil if hours.empty? or minutes.empty?
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
      return nil if url.empty?
      link_f = make_data_field('856', '4', '0', url)
      append_subfield link_f, 'y', message
      return link_f
    end

    def make_img_link(title, cover, thumb)
      return nil if cover.empty? or thumb.empty?
      img_f = make_data_field('856', '4', '2', cover)
      append_subfield img_f, 'y', "<img class=\"scl_mwthumb\" src=\"#{thumb}\" alt=\"Artwork for this title - #{title}\" />"
      return img_f
    end

  	def append_subfield(field, code, value)
  		return field.append(MARC::Subfield.new(code, value))
  	end

    def normalize_author(author)
  		return author if author.empty?
  		author   = author.split(',')[0]
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

    ##
    # Quickly turn 325645 {kb} into 318 {mb} etc.

    def kb_to_mb(size)
      return (size.to_f / 1024).to_i.to_s
    end

end
