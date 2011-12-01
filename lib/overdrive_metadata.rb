require 'marc'
require 'spreadsheet'

# Class to generate marc records from Overdrive provided metadata spreadsheet
# Usage:
# # Remove the header of the Overdrive spreadsheet and save it as .xls (not xml)
# require 'overdrive_metadata'
# records = OverdriveMetadata.new('spreadsheets/111111.xls')
# puts "R: " + records.size.to_s # print number of records generated to console
# w = MARC::Writer.new('generated.mrc')
# records.each do |r|
#   begin
#     w.write r
#   rescue
#     puts "FAILED: " + r['245']['a']
#   end
# end
# w.close

class OverdriveMetadata
    VERSION = '1.0.2'

    attr_reader :records

    DEF_FORMAT = 'eBook'
    OD_ORG     = 'OverDrive, Inc.'
    OD_URL     = 'http://www.overdrive.com'
    AGENCY     = 'JTH' # Sonoma County Library
    ACCESS     = 'Mode of access: World Wide Web.'
    URL_MSG    = 'Click to download this resource.'
    DISCLAIM   = 'Record generated from Overdrive metadata spreadsheet.'
    READ_ERR   = 'Error, close file, check file path or try resaving file as .xls (not xml)'

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
        rescue Exception => ex
            raise READ_ERR
        end
        @records  = []
        @count    = 0
        map
    end

    def map
        @metadata.each do |row|
            @records << create_record(row)
        end
      @records.compact!
      merge_by_content_url
    end

    def create_record(data)
      @count += 1
      field = package_data(data)
      r = field[:format].match(/#{DEF_FORMAT}/) ? EBook.new : EAudioBook.new
      begin
        r.make_control_field('001', field[:oclc])
        r.make_006
        r.make_007
        r.make_fixed_field(field[:year], field[:month], field[:day])
        r.make_data_field('020', ' ', ' ', {'a' => field[:isbn] + ' ' + r.isbn}) unless field[:isbn].empty?
        r.make_data_field('037', ' ', ' ', {'b' => OD_ORG, 'n' => OD_URL})
        r.make_data_field('040', ' ', ' ', {'a' => AGENCY, 'c' => AGENCY})
        r.make_data_field('100', '1', ' ', {'a' => normalize_author(field[:author])})
        r.make_title(field[:title], field[:author])
        r.make_publication(field[:place], field[:publisher], field[:year])
        r.make_physical(field[:hours], field[:minutes])
        r.make_data_field('306', ' ', ' ', {'a' => field[:hours] + field[:minutes] + field[:seconds]})
        r.make_data_field('538', ' ', ' ', {'a' => ACCESS})
        r.make_data_field('538', ' ', ' ', {'a' => 'Requires ' + field[:requires] + '.'})  
        r.make_data_field('500', ' ', ' ', {'a' => "#{field[:format]} (file size: #{field[:filesize]} MB)."})
        r.make_data_field('511', '0', ' ', {'a' => "Read by #{field[:reader]}."}) unless field[:reader].empty? 
        r.make_data_field('520', ' ', ' ', {'a' => field[:summary]}) unless field[:summary].match(/^#+$/)
        r.make_data_field('500', ' ', ' ', {'a' => "Title from: #{field[:title_src]}."})
        r.make_data_field('500', ' ', ' ', {'a' => 'Unabridged.'}) if r.is_a? EAudioBook
        r.make_data_field('500', ' ', ' ', {'a' => "Duration: #{field[:hours]} hr., #{field[:minutes]} min."}) if r.is_a? EAudioBook
        field[:subjects].each { |s| r.make_data_field('655', ' ', '7', {'a' => clean_string(s).strip + '.', '2' => 'local'}) }
        r.make_data_field('655', ' ', '7', {'a' => r.subject, '2' => 'local'})
        r.make_data_field('700', '1', ' ', {'a' => normalize_author(field[:reader])})
        r.make_data_field('856', '4', '0', {'u' => field[:download], 'y' => URL_MSG})
        r.make_data_field('856', '4', '0', {'u' => field[:excerpt], 'y' => "Excerpt (#{field[:format]})."})
        r.make_data_field('856', '4', '2', {'u' => field[:cover], 'y' => "<img class=\"scl_mwthumb\" src=\"#{field[:thumb]}\" alt=\"Artwork for this title - #{field[:title].gsub(/[^A-Za-z ]/, '')}\" />"})
        r.make_data_field('907', ' ', ' ', {'a' => 'ER'})
        r.make_data_field('991', ' ', ' ', {'a' => DISCLAIM})
        return r.record
      rescue Exception => ex
        puts @count.to_s + ': ' + "#{ex.message}\n" + ex.backtrace[0..2].join("\n")
        nil
      end
    end

    def package_data(data)
      values = {}
      values[:isbn]             = data[HEADERS[:isbn]]
      values[:date]             = data[HEADERS[:date]]
      values[:place]            = data[HEADERS[:place]]
      values[:publisher]        = data[HEADERS[:publisher]]
      values[:month]            = ''
      values[:day]              = ''
      if values[:date].match(/\d{1,2}\/\d{1,2}\/\d{4}/)
        month, day, year        = values[:date].split '/'
        values[:month]          = month
        values[:day]            = day
      end
      values[:year]             = values[:date].match(/\d{4}/).to_s # Fall-back
      values[:time]             = data[HEADERS[:time]]
      hr, mn, sc                = values[:time].split ':'
      values[:hours]            = hr ? hr : ''
      values[:minutes]          = mn ? mn : ''
      values[:seconds]          = sc ? sc : ''
      values[:author]           = clean_string data[HEADERS[:author]]
      values[:title]            = clean_string data[HEADERS[:title]]
      values[:title_src]        = data[HEADERS[:title_src]]
      values[:reader]           = clean_string data[HEADERS[:reader]]
      values[:requires]         = data[HEADERS[:requires]]
      values[:format]           = data[HEADERS[:format]]
      values[:filesize]         = kb_to_mb(data[HEADERS[:filesize]])
      values[:summary]          = clean_string data[HEADERS[:summary]]
      values[:subjects]         = data[HEADERS[:subjects]].split(',') rescue []
      values[:download]         = data[HEADERS[:download]]
      values[:excerpt]          = data[HEADERS[:excerpt]]
      values[:thumb]            = data[HEADERS[:thumb]]
      values[:cover]            = data[HEADERS[:cover]]
      values[:oclc]             = data[HEADERS[:oclc]].to_s.empty? ? 'ovr' + make_id(values[:download]) : 'ocn' + data[HEADERS[:oclc]]
      values.each { |k, v| values[k] = '' if v.nil? }
      return values
    end

    def merge_by_content_url
        puts 'Merging (may take a while on large record sets) ...'
        content_url = Hash.new(0)
        @records.each do |record|
          content_url[record['856']['u']] += 1
        end
        content_url.delete_if { |k,v| v < 2 }
        content_url.keys.each do |url|
          rcds = @records.find_all { |r| r['856']['u'] == url }
          raise 'Found invalid number of duplicate records: ' + url unless rcds.size == 2
          file_note = rcds[1].find { |f| f.tag == '500' and f['a'] =~ /OverDrive (WMA|MP3) Audiobook/ }
          excerpt   = rcds[1].find { |f| f.tag == '856' and f['y'] =~ /Excerpt/ }
          if file_note and excerpt
            begin
              rcds[0].fields.insert(rcds[0].fields.index { |f| f.tag == '500' }, file_note)
              rcds[0].fields.insert(rcds[0].fields.index { |f| f.tag == '856' and f['y'] =~ /Excerpt/ }, excerpt)
              @records.delete rcds[1]
            rescue Exception => ex
              puts url + ': ' + 'failed to merge'
            end
          end
        end
        @records
    end

    def make_id(id_string)
        return id_string[-9..-1].gsub(/\W/, '')
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

    def clean_string(input_str)
        return input_str.gsub(/&lt;.*&gt;/, '').gsub(/&quot;/, '"').gsub(/&apos;/, "'").gsub(/&#160;/, '').gsub(/&#235;/, 'e').gsub(/<\/?[^>]*>/, '').gsub(/\s{2}+/, ' ').strip rescue ''
    end

    # Quickly turn 325645 {kb} into 318 {mb} etc. + 1 so not 0

    def kb_to_mb(size)
      return (size.to_f / 1024 + 1).to_i.to_s
    end

    class ERecord

      GMD      = '[electronic resource]'
      DATE_ERR = 'Date information not present for fixed field'
      FIXF_ERR = 'Invalid fixed field created'
      TITL_ERR = 'Title data is missing for record'

      attr_reader :record

      def initialize
        @record = MARC::Record.new
        @ldr = record.leader
        @ldr[5]  = 'n'
        @ldr[7]  = 'm'
        @ldr[17] = 'M'
        @ldr[18] = 'a'
        fixed_field = ''
      end

      def make_control_field(tag, value)
        return nil if value.empty?
        @record.append MARC::ControlField.new(tag, value)
      end

      def make_data_field(tag, ind1, ind2, subfields)
        s = []
        subfields.each do |k,v|
          return nil if v.nil? or v.empty?
          s << MARC::Subfield.new(k, v)
        end
        @record.append MARC::DataField.new(tag, ind1, ind2, *s)
      end

      def make_fixed_field(year, month, day)
        raise DATE_ERR if year.empty?
        fixed_field = @fixed_field
        unless month.empty? and day.empty?
          month = '0' + month if month.length == 1
          day   = '0' + day if day.length == 1
          fixed_field[0..5] = year[2..3] + month + day
          fixed_field[7..10] = year
        else
          fixed_field[7..10] = year
        end
        raise FIXF_ERR unless fixed_field.length == 40
        make_control_field('008', fixed_field)
      end

      def make_title(title, sor)
        raise TITL_ERR if title.empty?
          t_ind1  = sor.empty? ? '0' : '1'
          t_ind2  = non_filing_characters title
          subfields = {}
          subfields['a'] = title
          subfields['h'] = GMD + ' /'
          unless sor.empty?
            value = sor[-1] == '.' ? "by #{sor}" : "by #{sor}."
            subfields['c'] = value
          else
            subfields['h'].gsub!(/\s+\/$/, '.')
          end
          make_data_field('245', t_ind1, t_ind2, subfields)
      end

      def make_publication(place, publisher, year)
        return nil if place.empty? or publisher.empty? or year.empty?
        make_data_field('260', ' ', ' ', {'a' => "#{place} :", 'b' => "#{publisher},", 'c' => "#{year}."})
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

    class EBook < ERecord

      attr_reader :isbn, :subject
        
      def initialize
          super
          @ldr[6]  = 'a'
          @fixed_field = '      s        xxu|||| s      0||| eng d'
          @isbn = '(electronic bk. : OverDrive Electronic Book)'
          @subject = 'Downloadable ebooks.'
          # set leader
      end

      def make_006
        make_control_field('006', 'm        d        ')
      end

      def make_007
        make_control_field('007', 'cr nnu---|||||')
      end

      def make_physical(*args)
        make_data_field('300', ' ', ' ', {'a' => "1 online resource."})
      end
      
    end

    class EAudioBook < ERecord

      attr_reader :isbn, :subject

      def initialize
        super
        @ldr[6]  = 'i'
        @fixed_field = '      s        xxunnnn s           eng d'
        @isbn = '(sound recording : OverDrive Audio Book)'
        @subject = 'Downloadable audiobooks.'
        # set leader
      end

      def make_006
        make_control_field('006', 'm        h        ')
      end

      def make_007
        make_control_field('007', 'sz usnnnnnnned')
        make_control_field('007', 'cr nna   |||||')
      end

      def make_physical(hours, minutes)
        return nil if hours.empty? or minutes.empty?
        make_data_field('300', ' ', ' ', {'a' => "1 sound file (ca. #{hours} hr., #{minutes} min.) :", 'b' => 'digital.'})
      end
      
    end

end