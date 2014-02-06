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
@filename_target = 'import_outlets_update.sql'  # target file
@sql=[]


def write_to_sqlfile(w)
  File.open(@filename_target, 'w') {|file| file.write(w)}
end


def build_command(data)
  # UPDATE Customers SET ContactName='Alfred Schmidt', City='Hamburg' WHERE CustomerName='Alfreds Futterkiste';
  
  return "UPDATE outlets SET temp_menu_images = \"#{data[:image]}\" WHERE slug = \"#{data[:slug]}\""
end


Spreadsheet.client_encoding = 'UTF-8'
book = Spreadsheet.open filename
sheet3 = book.worksheet 2


puts "[#{DateTime.now}] Start..... "
sheet3.each_with_index do |row,i|
  next if i==0
  
  data = {}
  data[:slug]  = row[1]
  data[:image] = row[2]
  

  @sql << build_command(data)

end
#puts "last: #{@sql.join("\n")}"
write_to_sqlfile(@sql.join(";\n"))

puts " [#{DateTime.now}] Generate [UPDATE]...... file \"#{@filename_target}\" "