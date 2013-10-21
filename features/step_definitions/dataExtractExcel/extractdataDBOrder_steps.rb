require 'win32ole'
require 'watir'

def extractdataexcelDBOrder


  #extracting saved data in excel for comparision to DB for dbo.Order

  @rows = @rows.to_s
  range1 = "A"+@rows
  range2 = "D"+@rows


  @excel = WIN32OLE::new("excel.Application")
  wrkbook = @excel.Workbooks.Open('G:\SSAutomation.xlsx')
  wrksheet = wrkbook.worksheets(1)
  wrksheet.select
  @arr = Hash.new
  @arr = wrksheet.Range("#{range1}:#{range2}").value
  wrkbook.close
  @excel.quit
  @excel.visible = 0

  @arr = @arr.flatten


  @ordnum = @arr[0].to_s
  @ordnum = @ordnum.chomp
  puts @ordnum
  @consultantCID = @arr[1].to_i
  @OrderStatus = @arr[2]
  @OrderTypeID = @arr[3].to_i


  @rows = @rows.to_i



end