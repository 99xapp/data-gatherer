version = '20.10.5'
vDetail = 'Vs 1.0'





require 'watir'
require 'rubyXL'
require 'rubyXL/convenience_methods'

#Add method to freeze first column
class RubyXL::Worksheet
  def add_frozen_split(row, column)
    worksheetview = RubyXL::WorksheetView.new
    worksheetview.pane = RubyXL::Pane.new(:top_left_cell => RubyXL::Reference.new(row,column),
    :y_split => row,
    :x_split => column,
    :state => 'frozenSplit',
    :activePane => 'bottomRight')
    worksheetviews = RubyXL::WorksheetViews.new
    worksheetviews << worksheetview
    self.sheet_views = worksheetviews
  end
end

puts
puts
puts '--------------------------------'
puts '---------INV Explorer ----------'
puts


userName = "your user name"

StockArray = ['SON','GM','F','TSLA']
#or, a stock list could be read from the computer as a text file
#  File.foreach('path to the file') do |line|
#    StockArray.push(line.chomp)
#  end


										 
Watir.default_timeout = 60 # change as needed


#-------------------- ^^^^ Set Parameters ^^^^ ---------------------




#------ Open Browser ------

b = Watir::Browser.new

#------------------------

# colors (I may use these later)

white        = 'ffffff'
yellow       = 'ffff66'
lightyellow  = 'ffffaa'
blue         = '99aaff'
lightblue    = 'ccddff'
red          = 'ff6666'
lightred     = 'ffbbbb'
green        = '66ff66'
lightgreen   = 'bbffbb'
orange       = 'ffccaa' #light orange
gray         = '666666'


rr = 0 #current row
cc = 0 #current column

#i = 0 #general purpose counter


#------ Acquire site ------

b.goto('https://www.investing.com')


#---------------------

#---------------------------

sleep(initDelay)   # not used right now

#---------------------------


	
sTime = Time.new.to_s  #start time
sTime = sTime[11,8]
 

# ------ Start new Book ------

resultsBook = RubyXL::Workbook.new
rs1 = resultsBook.worksheets[0]
rs1.add_frozen_split(0,1)


## ------------------------------------------ 


# ------ Results Location -----------

# ------ Set Filename and Path ------

time = Time.new
y = time.year.to_s
m = time.month.to_s
d = time.day.to_s

  fileName = 'Stocks ' + ' ' + y + '.' + m + '.' + d

  savePath = '/Users/' + userName + '/Dropbox/------/------/------/' + fileName + '.xlsx'
  #Example savePath = '/Users/' + userName + '/Dropbox/Dempsey/' + filename + '.xlsx'
  puts 'Save as ' + savePath


  #=========== Heading =============

  rs1.add_cell(0, 0,'STOCK')
  rs1[0][0].change_fill(yellow)
  rs1.add_cell(0, 1,'PRICE')
  rs1[0][1].change_fill(yellow)
  rs1.add_cell(0, 2,'CHANGE')
  rs1[0][2].change_fill(yellow)
  rs1.add_cell(0, 3,'%')
  rs1[0][3].change_fill(yellow)
  
  
  # ============ Search Loop ============
  
StockArray.each do |stock|
  
  b.div(class: 'searchBoxContainer').text_field(class: 'searchText').set stock
  #b.text_field(class: "searchText").set "SON"  also works
	b.send_keys :enter


  # select the top one
  
  div1 = b.div(class: 'searchSectionMain').wait_until(message: 'div1 ?') { |el| el.present? }
  sLink = div1.link(class: ['js-inner-all-results-quote-item','row'])
  sLink.click
 
  # now read the data
 
 # deeper specification seems to increase speed
 # ie, sCost can be found with sCostRaw = b.span(id: 'last_last').wait_until(message: 'Cost ?') { |el| el.present? }
 # but giving a better spec helps Watir find the element with less searching - I think ...
 
 	sDiv = b.div(class: 'overViewBox').div(id: 'quotes_summary_current_data').div(class: 'current-data').div(class: 'inlineblock').div(class: 'top')
 
  sCostRaw = sDiv.span(id: 'last_last').wait_until(message: 'Cost ?') { |el| el.present? }
  sCost = sCostRaw.text
  puts sCost
 
  sChngRaw = sDiv.span(class: 'arial_20')
  sChng = sChngRaw.text
  puts sChng
 
  pCentRaw = sDiv.span(class: 'parentheses')
  pCent = pCentRaw.text
  puts pCent


#  sChngRaw = b.span(class: ['arial_20', 'greenFont', 'pid-20726-pc'])
#  sChng = sChngRaw.text
#  puts sChng
 
#  pCentRaw = b.span(class: ['arial_20', 'greenFont', 'pid-20726-pcp', 'parentheses'])
#  pCent = pCentRaw.text
#  puts pCent
 
 	rr += 1  #advance to next row
  # Write to workbook
  rs1.add_cell(rr,0,stock)
  rs1.add_cell(rr,1,sCost)
  rs1.add_cell(rr,2,sChng)
  rs1.add_cell(rr,3,pCent)

 
 # Write the Book to the computer
 resultsBook.write savePath


end  # of search loop


  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
