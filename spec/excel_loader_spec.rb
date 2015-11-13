require Dir.getwd + '/lib/excel_loader'

describe ExcelLoader do

  PATH = Dir.getwd + "/spec/fixtures.xls"
  
  describe "process" do

    it "should yield each row as a hash" do
      ExcelLoader.process(PATH) do |row|
        row.class.should == Hash
        row[:name].should == "Mark Twain" if row[:id] == 2
      end
    end

  end
  
  describe "file_to_array" do
    
    before(:all) do
      @h = ExcelLoader.file_to_array(PATH)
    end
    
    it "should take a file name and return an array of hashes" do
      @h.class.should == Array
      @h[0].class.should == Hash
    end
    
    it "should take the value of the formula if the cell is a formula" do
      @h[2][:id].should == 3
    end
    
    it "should set the value to nil if a cell is blank" do
      @h[2].keys.include?(:residence).should == true
      @h[2][:residence].should == nil
    end
    
    it "should only serialize up to the last row" do
      @h.size.should == 4
    end
    
    it "should only take data from columns that have a header" do
      @h[1].size.should == 3
    end
    
    it "shouldn't worry about bold or italics" do
      @h[1][:id].should == 2
      @h[1][:name].should == "Mark Twain"
    end
    
    it "should take an optional zero-based sheet index parameter" do
      ExcelLoader.file_to_array(PATH, 1).should == [{:id=>1}]
    end
    
    it "should return an empty array if there is no data after the first row" do
      ExcelLoader.file_to_array(PATH, 2).should == []
    end
    
    it "should return an empty array if there is no data in the first row" do
      ExcelLoader.file_to_array(PATH, 3).should == []
    end
    
    it 'works with .xlsx files' do
      xlsx = PATH.gsub(".xls", ".xlsx")
      ExcelLoader.file_to_array(PATH, 0).should == ExcelLoader.file_to_array(xlsx, 0)
      ExcelLoader.file_to_array(PATH, 1).should == ExcelLoader.file_to_array(xlsx, 1)
      ExcelLoader.file_to_array(PATH, 2).should == ExcelLoader.file_to_array(xlsx, 2)
      ExcelLoader.file_to_array(PATH, 3).should == ExcelLoader.file_to_array(xlsx, 3)
    end
    
  end
  
  before(:all) do
    @data = [
      {:id => 1, :name => "Stephen King"},
      {:id => 2, :name => "Tom Clancy"},
      {:id => 3, :name => "Umberto Eco"}
    ]
  end
  
  describe "hashes_to_arrays" do
    
    it "should take an array of hashes and return an array of arrays" do
      arr = ExcelLoader.send(:hashes_to_arrays, @data)
      arr[1].size.should == 2
      arr[1].should include(1)
      arr[1].should include("Stephen King")
    end
    
    it "should stringify the keys and put them in the first entry" do
      arr = ExcelLoader.send(:hashes_to_arrays, @data)
      arr[0].size.should == 2
      arr[0].should include("id")
      arr[0].should include("name")
    end
    
    it "should take an optional mapping parameter that orders the elements" do
      arr = ExcelLoader.send(:hashes_to_arrays, @data, [:id, :name])
      arr.should == [["id", "name"], [1, "Stephen King"], [2, "Tom Clancy"], [3, "Umberto Eco"]]
    end
    
  end
  
  describe "array_to_file" do
    
    it "should take an array of hashes and return a path string" do
      path = ExcelLoader.array_to_file(@data)
      path.class.should == String
      File.delete(path)
    end
    
    it "should take an array of arrays also " do
      arr = [["id", "name"], [1, "Stephen King"], [2, "Tom Clancy"], [3, "Umberto Eco"]]
      path = ExcelLoader.array_to_file(arr)
      File.exists?(path).should == true
      File.delete(path)
    end
    
    it "should raise an error if anything but an array of 
      hashes or an array of arrays is passed" do
      err = "First parameter must be an array of hashes or an array of arrays"
      lambda { ExcelLoader.array_to_file([:blah]) }.should raise_error(err)
      lambda { ExcelLoader.array_to_file([[:id], {:id=>1}]) }.should raise_error(err)
      lambda { ExcelLoader.array_to_file("hi") }.should raise_error
    end
    
    it "should take an optional path parameter and set the file name to that" do
      path = ExcelLoader.array_to_file(@data, nil, "data.xls")
      File.exists?(path).should == true
      File.delete(path)
    end
    
    it 'works with .xlsx files' do
      path = ExcelLoader.array_to_file(@data, nil, "data.xlsx")
      File.exists?(path).should == true
      ExcelLoader.file_to_array("data.xlsx").should == @data
      File.delete(path)
    end
    
  end
  
  describe "headers" do
    
    before(:all) do
      @arr = ExcelLoader::headers(PATH)
    end

    it "should return an array" do
      @arr.class.should == Array
    end
    
    it "should be ordered by column index" do
      @arr.should == [:id, :name, :residence]
    end

  end
  
end