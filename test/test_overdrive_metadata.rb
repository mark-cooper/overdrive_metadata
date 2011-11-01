require "shoulda"
require "overdrive_metadata"

class TestOverdriveMetadata < Test::Unit::TestCase
  
  context "Creating Overdrive records" do
  
  	setup do 
  		@o = OverdriveMetadata.new('raw/test.xls')
  	end 
  	
  	should "Make a control field" do 
  		assert_equal '006', @o.make_control_field('006', 'm        h        ').tag
  		assert_equal 'm        h        ', @o.make_control_field('006', 'm        h        ').value
  	end

  	should "Make a nil control field" do 
  		assert_equal nil, @o.make_control_field('006', '')
  	end

  	should "Make a data field" do 
  		assert_equal '037', @o.make_data_field('037', '', '', 'OverDrive, Inc.', 'b').tag
  		assert_equal 'OverDrive, Inc.', @o.make_data_field('037', '', '', 'OverDrive, Inc.', 'b').value
  		assert_equal 'OverDrive, Inc.', @o.make_data_field('037', '', '', 'OverDrive, Inc.', 'b')['b']
  	end

  	should "Make a nil data field" do 
  		assert_equal nil, @o.make_data_field('245', '1', '4', '')
  	end

  	should "Make a fixed field" do 
  		assert_equal '      s2011    xxunnnn s           eng d', @o.make_fixed_field('Jul  5 2011 12:00AM')
  		assert_equal '110726s2011    xxunnnn s           eng d', @o.make_fixed_field('07/26/2011')
  	end

  end

end
