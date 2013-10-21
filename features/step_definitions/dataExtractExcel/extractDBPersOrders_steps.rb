require 'win32ole'
require 'watir'



def extractDBpersorders


  @rows = @rows.to_s
  range1 = "A"+@rows
  range2 = "R"+@rows


  excel = WIN32OLE::new("excel.Application")
  wrkbook = excel.Workbooks.Open('G:\SSAutomationPulse.xlsx')
  wrksheet = wrkbook.worksheets(2)
  wrksheet.select
  @arr = Hash.new
  @arr = wrksheet.Range("#{range1}:#{range2}").value
  wrkbook.close
  excel.quit
  @arr = @arr.flatten


  @product = @arr[0]
  @OrderNumber = @arr[1].to_s
  @OrderNumber = @OrderNumber.chomp
  @job = @arr[2]
  @style = @arr[3]
  @shortstyle = @arr[4]
  @hoop = @arr[5]
  @textline1 = @arr[6]
  @textline2 = @arr[7]
  @thread = @arr[8]
  @status = @arr[9]
  @designStyle = @arr[10]
  @designColors = @arr[11]


  @rows = @rows.to_i

end