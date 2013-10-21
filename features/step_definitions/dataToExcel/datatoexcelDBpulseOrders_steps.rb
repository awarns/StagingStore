require 'rubygems'
require 'watir'


def datatoexcelDBpulseorders


  if @personalizationtypeid > '1'

    @excel = WIN32OLE::new("excel.Application")
    wrkbook = @excel.Workbooks.Open('G:\SSAutomationPulse.xlsx')
    wrksheet = wrkbook.worksheets(2)
    wrksheet.select



    wrksheet.Cells(@rows, "A").value = @itemalias
    wrksheet.Cells(@rows, "B").value = @ordnum

    wrksheet.Cells(@rows, "D").value = @embroiderfontstyle

    @shortstyle = @embroiderfontstyle.split (/- */)
    @shortstyle = @shortstyle[0]
    @hoop = @shortstyle[1]
    wrksheet.Cells(@rows, "E").value = @shortstyle
    wrksheet.Cells(@rows, "F").value = @hoop
    wrksheet.Cells(@rows, "G").value = @text_line_1
    wrksheet.Cells(@rows, "H").value = @text_line_2
    wrksheet.Cells(@rows, "I").value = @font_color
    wrksheet.Cells(@rows, "J").value = 'processed'
    wrksheet.Cells(@rows, "K").value = @PulseDesignStyle
    wrksheet.Cells(@rows, "L").value = @PulseDesignColors


    wrkbook.save
    excel.quit




    @option1 = nil
    @option2 = nil
    @option3 = nil
    @option4 = nil
    @option5 = nil
    @textbox1 = nil
    @textbox2 = nil
    @print = nil
    @personalization = nil
    @personalizationtypeid = nil
    @designstyle = nil
    @embroiderfontstyle = nil
    @numberoflines = 0



  end


  @option1 = nil
  @option2 = nil
  @option3 = nil
  @option4 = nil
  @option5 = nil
  @textbox1 = nil
  @textbox2 = nil
  @print = nil
  @personalization = nil
  @personalizationtypeid = nil
  @designstyle = nil
  @embroiderfontstyle = nil
  @numberoflines = 0




end