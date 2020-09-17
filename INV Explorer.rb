version = '20.915'
vDetail = 'Vs 0.0'

# 3XT will be able to run multiple data sets consecutively




# WORK ON TIMEOUT EXCEPTION !!
#

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
puts '-------INV Explorer -------'
puts

###############################################################################

#--------------------- Set Parameters Here --------------------------

#userName = "williamhaynes" #iMac at Dovedale & MBA
userName = "alhaynes"      #iMac at lake


stockArray = ['SON']
nStocks = stockArray.size

stock = 'SON' # not going to use an array at first

compDay = 'x' # day  
#     or
compDate = 'x' # entire date  
# only set 1 or the other - not both !
# data is a String !
if compDay.upcase != 'X' && compDate.upcase != 'X'
	puts 'Error in Comp Setting !'
end

# ---------------------------------------------------------------------------
if compDay.upcase == 'X' && compDate.upcase == 'X' #--- Don't Set Here ! ---#
	comp = 0
else
	comp = 1
end	
#-----------------------------------------------


## ------ Set Delays ------

delayTime = 3600 * 0  # SET DELAY HERE 3600 = 1 hour -----------------
sleep(delayTime)

initDelay     = 2    # for first search
loadDelay     = 5    # each pn load
loopDelay1    = 3    # for each parts loop
loopDelay2    = 3    # for each parts loop
dlrTableDelay = 1
pnDelay       = 5    # If pn isn't fresh 
newBookDelay  = 5  # between workbooks (not 1st book, though)
										 
Watir.default_timeout = 30 # change as needed

#-------------------- ^^^^ Set Parameters ^^^^ ---------------------


##############################################################################



#------ Open Browser ------

b = Watir::Browser.new

#------------------------

# colors (I'll use these later)

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

i = 0 #general purpose counter


#------ Sign In ------

b.goto('https://www.investing.com')



#---------------------

#---------------------------

sleep(initDelay)   # not used right now

#---------------------------


#Start the Search
	
  sTime = Time.new.to_s
  sTime = sTime[11,8]
 

	# Start new Book
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

  fileName = 'Prices ' + ' ' + y + '.' + m + '.' + d

  savePath = '/Users/' + userName + '/Dropbox/Dempsey/Stock ' + fileName + '.xlsx'
  puts 'Save as ' + savePath


  #=========== Heading =============

  rs1.add_cell(0, 0,'STOCK')
  rs1[0][0].change_fill(yellow)
  rs1.add_cell(0, 1,'PRICE')
  rs1[0][1].change_fill(yellow)
  
  
  #------ Search -------------
  
  # faster !  this does work, but slow
  b.div(class: 'searchBoxContainer').text_field(class: 'searchText').set 'SON'
  #b.text_field(class: "searchText").set "SON"  also works
	b.send_keys :enter


 # select SON - the top one
  
 div1 = b.div(class: 'searchSectionMain').wait_until(message: 'div1 ?') { |el| el.present? }
 sLink = div1.link(class: ['js-inner-all-results-quote-item','row'])
 sLink.click
 
 # now read the data
 
 sCostRaw = b.span(id: 'last_last').wait_until(message: 'Cost ?') { |el| el.present? }
 sCost = sCostRaw.text
 puts sCost
 
 
 # Write to workbook
  rs1.add_cell(1,0,'SON')
  rs1.add_cell(1,1,sCost)

 
 # Write the Book to the computer
 resultsBook.write savePath








#----Unused ----
#  <div class="searchBoxContainer topBarSearch topBarInputSelected">
#	    <input autocomplete="off" type="text" class="searchText arial_12 lightgrayFont js-main-search-bar" value="" placeholder="Search the website...">
#		<label class="searchGlassIcon js-magnifying-glass-icon">&nbsp;</label>
#		<i class="cssSpinner"></i>
#	</div>

  	#b.button(class: "searchGlassIcon").click didn't find it ...

  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
