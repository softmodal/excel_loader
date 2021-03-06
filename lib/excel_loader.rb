require 'rubygems'
require 'spreadsheet'
require 'rubyXL'

module ExcelLoader
  # Dumps an ExcelLoader file into an array of hashes.  The file must be laid out like a table.
  # The first row must contain header information and each column must have a header.  There
  # cannot be empty columns to the left of the table.
  #
  # ==== Examples
  #   example.xls looks like this:
  #   -----------------------------------------
  #   | id  | name                | residence |
  #   -----------------------------------------
  #   |  1  | Tom Clancy          | Maryland  |
  #   -----------------------------------------
  #   |  2  | Umberto Eco         | Italy     |
  #   -----------------------------------------
  #   |  3  | Patrick O'Brien     | England   |
  #   -----------------------------------------
  #
  #   objects = ExcelLoader.file_to_array("example.xls")
  #   objects == [{:id => 1, :name => "Tom Clancy", :residence => "Maryland"},
  #                   {:id => 2, :name => "Umberto Eco", :residence => "Italy"},
  #                   {:id => 3, :name => "Patrick O'Brien", :residence => "England"}]
  #
  # ==== Parameters
  # path<String>::  Path to the ExcelLoader file.
  # index<Integer>:: Zero-based worksheet index.
  #     Defaults to +0+
  #
  # ==== Returns
  # Array:: An array of hashes representing the rows in the file.
  def self.file_to_array(path, index=0)
    rows = []
    headers = self.headers(path, index)
    i = 0
    self.process(path, index) do |row|
      i += 1
      next if i == 1
      break if row == {} or (row.values.map { |v| v.to_s.strip } - [""]).empty?
      headers.each { |header| row[header] = nil if row[header].nil? }
      rows.push(row)
    end
    rows
  end
  
  def self.workbook(path)
    path.to_s.match(/\.xlsx/i) ? RubyXL::Parser.parse(path) : Spreadsheet.open(path)
  end
  
  def self.sheet(path, index=0)
    book = workbook(path)
    book.respond_to?(:worksheet) ? book.worksheet(index) : book[index]
  end
  
  def self.process(path, index=0)
    headers = self.headers(path, index)
    sheet(path, index).each do |row|
      obj = {}
      cells(row).each_with_index do |cell, i|
        if headers[i]
          obj[headers[i]] = cell.respond_to?(:value) ? cell.value : cell
        end
      end
      yield obj
    end
  end

  # Dumps an array into an ExcelLoader file.  The file will be laid out like a table.  
  # If an array of arrays is passed then the rows will correspond with the 
  # arrays.  If an array of hashes is passed then the first row will contain 
  # header information, composed of the keys in each hash.  
  #
  # If a mapping array is passed with an array of hashes then the columns 
  # will placed in order.  Otherwise, the columns will be entered at random.
  #
  # If a path is passed then the file will be saved there.  Otherwise, the
  # the file is saved with a random name in the current directory.
  #
  # ==== Examples
  #   objects = [{:id => 1, :name => "Tom Clancy", :residence => "Maryland"},
  #                  {:id => 2, :name => "Umberto Eco", :residence => "Italy"},
  #                  {:id => 3, :name => "Patrick O'Brien", :residence => "England"}]
  #
  #   ExcelLoader.array_to_file(objects, [:id, :name, :residence], "example.xls")
  #
  #   Creates example.xls which look like this:O
  #   -----------------------------------------
  #   | id  | name                | residence |
  #   -----------------------------------------
  #   |  1  | Tom Clancy          | Maryland  |
  #   -----------------------------------------
  #   |  2  | Umberto Eco         | Italy     |
  #   -----------------------------------------
  #   |  3  | Patrick O'Brien     | England   |
  #   -----------------------------------------
  #
  # ==== Parameters
  # arr<Array>::  Array of data.
  # mapping<Array>::  Array of symbols of mapping information.
  #     Defaults to +nil+
  # path<String>:: Path of the created ExcelLoader file.
  #     Defaults to +nil+
  #
  # ==== Returns
  # String:: The path of the created file.
  def self.array_to_file(arr, mapping=nil, path=nil)
    path ||= "#{rand(9999999999)}.xls"
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet(:name => 'Sheet1')
    if path.match(/\.xlsx$/i)
      book = RubyXL::Workbook.new
      sheet = book[0]
    end
    
    types = arr.map { |item| item.class }.uniq
    unless types == [Hash] or types == [Array]
      raise "First parameter must be an array of hashes or an array of arrays"
    end
    
    rows = types == [Hash] ? hashes_to_arrays(arr, mapping) : arr
    if path.match(/\.xlsx$/i)
      rows.each_with_index do |row, i|
        row.each_with_index do |val, j|
          sheet.add_cell(i, j, val)
        end
      end
    else
      rows.each_index do |index|
        sheet.row(index).concat(rows[index])
      end
    end
    book.write(path)
    path
  end
  
  def self.cells(row)
    return [] if !row
    row.respond_to?(:cells) ? row.cells : row
  end
  
  # Returns an array of the headers of this file
  def self.headers(path, index=0)
    sht = sheet(path, index)
    row = sht.respond_to?(:row) ? sht.row(0) : sht[0]
    cells = 
    arr = []
    cells(row).each do |cell|
      next if !cell
      val = cell.respond_to?(:value) ? cell.value : cell
      arr.push(val.to_s.intern)
    end
    arr
  end
  
  private
  def self.hashes_to_arrays(arr, mapping=nil)
    keys = mapping || arr[0].keys
    rows = [keys.map { |key| key.to_s }]
    arr.each do |item|
      row = []
      keys.each do |key|
        row.push(item[key])
      end
      rows.push(row)
    end
    rows
  end
  
end

