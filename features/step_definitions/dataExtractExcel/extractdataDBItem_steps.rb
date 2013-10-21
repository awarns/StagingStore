require 'win32ole'
require 'watir'


def extractdataexceldbitem


  #Extracting everthing on Sheet 2 of SS Automation Excel for dbo.item.
  #Formating properly to get ready to compare to Database


  @rows = @rows.to_s
  range1 = "A"+@rows
  range2 = "AI"+@rows


  @excel = WIN32OLE::new("excel.Application")
  wrkbook = @excel.Workbooks.Open('G:\SSAutomation.xlsx')
  wrksheet = wrkbook.worksheets(2)
  wrksheet.select
  @arr = Hash.new
  @arr = wrksheet.Range("#{range1}:#{range2}").value
  wrkbook.close
  @excel.quit
  @excel.visible = 0
  @arr = @arr.flatten


  @OrdersystemUUID = @arr[0]
  @personalizationtypeid = @arr[1].to_i
  @itemalias = @arr[2]
  @linenumber = @arr[3].to_i
  @lineseq = @arr[4].to_i
  @qty = @arr[5].to_i
  @barcode = @arr[6].to_i
  @embroidfontcolor = @arr[7]
  @embroidfontstyle = @arr[8]
  @embroidline1 = @arr[9]
  @embroidline2 = @arr[10]
  @embroidline3 = @arr[11]
  @iskitheader = @arr[12].to_s.downcase
  @iskitcomponent= @arr[13].to_s.downcase
  @kitheader = @arr[14]
  @textepression = @arr[15]
  @Kid1 = @arr[16]
  puts @Kid1
  @Kid1text = @arr[17]
  puts @Kid1text
  @Kid2 = @arr[18]
  puts @Kid2
  @Kid2Text = @arr[19]
  puts @Kid2Text
  @Kid3 = @arr[20]
  puts @Kid3
  @Kid3Text = @arr[21]
  puts @Kid3Text
  @Kid4 = @arr[22]
  puts @Kid4
  @Kid4Text = @arr[23]
  puts @Kid4Text
  @Kid5 = @arr[24]
  puts @Kid5
  @Kid5Text = @arr[25]
  puts @Kid5Text
  @Kid6 = @arr[26]
  puts @Kid6
  @Kid6Text = @arr[27]
  puts @Kid6Text
  @StationaryStyle = @arr[28]
  @ItemStatus = @arr[29]
  @DesignStyle = @arr[30]
  @DesignColor = @arr[31]
  @ParentItemID = @arr[32]
  @KidTextPrefix = @arr[33]
  @Price = @arr[34]


  @rows = @rows.to_i


end