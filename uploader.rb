# 
# This program will generate .sql file from .xls / .xlsx file , and ready to upload to db 
# avoid direct connection with db , to reduce missing field value , etc.
# 
# [created by Ary]
# 
# \, A, wrong field value, '.0'
require "roo"
require "spreadsheet"
require "nokogiri"


filename = 'source_data_from_entry_team.xls'  # source file
@filename_target = 'import_outlets.sql'  # target file
sheet_index = 9
@cols_num = 23   # total column start from index zero.
@col_notstrtype = ['enabled','city_id', 'neighbourhood_id', 'mall_id', 'latitude', 'longitude', 'price_id']  # column integer type



@col_notstrtype_id= []  # column integer type index


def generate2sql(w)
  File.open(@filename_target, 'w') {|file| file.write(w)}
end

def make_row(tmp, row, i)
  row.map {|w| 
    break if row.index(w) > @cols_num
    tmp << append_quote(w.to_s, row.index(w), i) << ","
  }
end

def append_quote(w,idx, rowindex)
  w.gsub!(/\"|\\"/i, "\'")
  @col_notstrtype_id << idx if @col_notstrtype.include?(w)
  return (@col_notstrtype_id.include?(idx) && (rowindex>0)) || (rowindex==0) ? "#{w}" : "\"#{w.strip}\""
end

def create_header(w)
  out = ''
  out << "INSERT INTO `outlets` "
  out << w.chomp(",\n")
  out << " VALUES "
end



# Main

Spreadsheet.client_encoding = 'UTF-8'
book = Spreadsheet.open filename
sheet = book.worksheet sheet_index


sql = ""

puts "[#{DateTime.now}] Start..... "
sheet.each_with_index do |row,i|
  
  tmp=""
  make_row(tmp, row, i)
  tmp.chomp!(",")
  
  sql += "(#{tmp}),\n"
  sql = create_header(sql.chomp(",")) if i == 0

end

# generate sql file
generate2sql(sql.chomp(",\n"))
puts " [#{DateTime.now}] Generate....... file \"#{@filename_target}\" "