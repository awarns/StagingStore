require 'win32ole'
require 'watir'


def datatoexcelDBOrder


  #formating items captured during order placement and sending them to excel for table dbo.Order

  #puts @ordnum
  #puts @consultantCID

  #puts @option1
  #puts @option2
  #puts @option3
  #puts @option4
  #puts @option5
  #puts @textbox1
  #puts @textbox2
  #puts @print
  #puts @personalization


  excel = WIN32OLE::new("excel.Application")
  wrkbook = excel.Workbooks.Open('G:\SSAutomation.xlsx')


  wrksheet = wrkbook.worksheets(1)
  wrksheet.select

  wrksheet.Cells(@rows, "A").value = @ordnum
  wrksheet.Cells(@rows, "B").value = @consultantCID
  wrksheet.Cells(@rows, "C").value = "Entered"
  wrksheet.Cells(@rows, "D").value = "16"

  wrkbook.save
  excel.quit


end









