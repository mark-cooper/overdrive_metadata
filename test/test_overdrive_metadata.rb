require "shoulda"
require "overdrive_metadata"

class TestOverdriveMetadata < Test::Unit::TestCase
  
  context "Creating Overdrive records" do
  
  	setup do 
  		@o = OverdriveMetadata.new('raw/test.xls')
  	end

    # Write some tests ...

  end

end
